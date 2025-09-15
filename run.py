import os
import yaml
from crewai import Agent, Task, Crew, Process
from docx import Document
import google.generativeai as genai
import logging
import time
import random
from pathlib import Path
from src.docx_preserve import (
    extract_text_units, replace_text_in_document, 
    calculate_optimal_batch_size, estimate_tokens, 
    TokenStats, log_token_stats
)


API_KEY = "AIzaSyAoiINUaIRTpBlcIIQZSdMsw7_VXa6S1c8"

if not API_KEY:
    raise ValueError("GOOGLE_API_KEY not found. Please set it in your .env file.")

genai.configure(api_key=API_KEY)
os.environ["GOOGLE_API_KEY"] = API_KEY
os.environ["GEMINI_API_KEY"] = API_KEY
os.environ["LLM_PROVIDER"] = "google"  
# Prevent accidental Vertex fallback via ADC
os.environ.pop("GOOGLE_APPLICATION_CREDENTIALS", None)
os.environ.pop("VERTEXAI_PROJECT", None)
os.environ.pop("VERTEXAI_LOCATION", None)



try:
    with open('config.yaml', 'r') as file:
        config = yaml.safe_load(file)
except FileNotFoundError:
    raise FileNotFoundError("config.yaml not found. Please ensure it is in the same directory.")



os.environ["OPENAI_MODEL_NAME"] = config['llm']['model_name']
os.environ["OPENAI_API_KEY"] = "EMPTY"



class TranslationAgents:
    def identification_agent(self):
        return Agent(
            role='Document Analyst',
            goal='Identify the document type (e.g., academic paper, technical report) and its primary language.',
            backstory='An expert in analyzing document structures and language patterns.',
            verbose=True,
            allow_delegation=False
        )

    def translator_agent(self):
        return Agent(
            role='Academic Language Translator',
            goal='Translate the entire provided text into the target language, ensuring the academic tone and context are preserved.',
            backstory='A polyglot and subject matter expert in academic translations, ensuring accuracy and nuance.',
            verbose=True,
            allow_delegation=False
        )

# --- 5. DEFINE TASKS ---
class TranslationTasks:
    def identify_task(self, agent, doc_content):
        return Task(
            description=f"""Analyze this document snippet to determine its type and original language.

            Content Snippet:
            ---
            {doc_content[:1000]}
            ---
            """,
            expected_output="A one-sentence summary confirming the document type and language.",
            agent=agent
        )

    def translate_task(self, agent, doc_content, target_language):
        return Task(
            description=f"""Translate the ENTIRE academic text provided below into {target_language}.
            You must translate everything accurately. Preserve the meaning, academic tone, and formatting like paragraphs and headings.
            
            Full Text to Translate:
            ---
            {doc_content}
            ---
            """,
            expected_output=f"The full and complete translated text in {target_language}, formatted with Markdown.",
            agent=agent,
            async_execution=False # Ensures this complex task is handled patiently
        )


# --- 6. SCRIPT EXECUTION ---
def run_crew():
    # Read the DOCX file
    def read_docx(file_path):
        try:
            doc = Document(file_path)
            # Find and remove the image placeholder if it exists
            full_text = [para.text for para in doc.paragraphs if "[Image of" not in para.text]
            return '\n'.join(full_text)
        except Exception as e:
            return f"Error reading document: {e}"

    input_path = config['paths']['input_file']
    output_path = config['paths']['output_file']
    Path(Path(output_path).parent).mkdir(parents=True, exist_ok=True)

    logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')
    logging.info("Reading input document: %s", input_path)

    document_content = read_docx(input_path)
    target_language = config['translation']['target_language']

    # Instantiate agents and tasks
    agents = TranslationAgents()
    tasks = TranslationTasks()

    # Build units for structure-preserving translation
    units, src_doc = extract_text_units(input_path)
    logging.info("Found %d text units for translation", len(units))

    # Agents
    identification_agent = agents.identification_agent()
    translator_agent = agents.translator_agent()

    # Identification task (context / prompt priming)
    task_identify = tasks.identify_task(identification_agent, document_content)

    # Calculate optimal batch size based on token limits
    max_tokens_per_batch = config['llm'].get('max_tokens_per_batch', 2000)
    batch_size = calculate_optimal_batch_size(units, max_tokens_per_batch)
    logging.info(f"Calculated optimal batch size: {batch_size} units")
    
    translated_map = []
    total_batches = (len(units) + batch_size - 1) // batch_size
    total_stats = TokenStats()
    
    for i in range(0, len(units), batch_size):
        batch = units[i:i+batch_size]
        batch_text = "\n\n--- UNIT BREAK ---\n\n".join(u.text for u in batch)
        
        # Calculate input tokens for this batch
        input_tokens = estimate_tokens(batch_text)
        
        task_translate = tasks.translate_task(translator_agent, batch_text, target_language)
        task_translate.context = [task_identify]
        crew = Crew(
            agents=[identification_agent, translator_agent],
            tasks=[task_translate],
            process=Process.sequential,
            verbose=False
        )
        
        batch_num = (i // batch_size) + 1
        logging.info(f"Translating batch {batch_num}/{total_batches} (units {i+1}-{min(i+batch_size, len(units))})")

        # Retry on transient provider errors (e.g., 503 overload)
        max_retries = 5
        attempt = 0
        translated_block = None
        batch_stats = TokenStats()
        
        while attempt < max_retries:
            try:
                translated_block = str(crew.kickoff())
                
                # Calculate output tokens and update stats
                output_tokens = estimate_tokens(translated_block)
                batch_stats.input_tokens = input_tokens
                batch_stats.output_tokens = output_tokens
                batch_stats.total_cost = (input_tokens + output_tokens) * 0.00001  # Rough cost estimate
                
                # Log batch statistics
                log_token_stats(batch_stats, batch_num, total_batches)
                
                # Update total stats
                total_stats.input_tokens += input_tokens
                total_stats.output_tokens += output_tokens
                total_stats.total_cost += batch_stats.total_cost
                
                break
            except Exception as e:
                attempt += 1
                if attempt >= max_retries:
                    raise
                backoff_seconds = min(30, (2 ** attempt) + random.uniform(0, 1))
                logging.warning(
                    "Translation batch failed (attempt %d/%d): %s. Retrying in %.1fs",
                    attempt, max_retries, e, backoff_seconds,
                )
                time.sleep(backoff_seconds)
        parts = [p.strip() for p in translated_block.split("--- UNIT BREAK ---")]
        if len(parts) != len(batch):
            # Fallback: try naive split by double newline
            parts = [p.strip() for p in translated_block.split("\n\n")]
        # Align best-effort
        for j, u in enumerate(batch):
            translated_text = parts[j] if j < len(parts) else translated_block
            translated_map.append((u, translated_text))

    # Log final statistics
    logging.info(f"Translation completed! Total stats - Input: {total_stats.input_tokens} tokens, "
                f"Output: {total_stats.output_tokens} tokens, Total cost: ${total_stats.total_cost:.4f}")
    
    # Replace text in a copy of the source doc to preserve layout/images
    out_doc = replace_text_in_document(src_doc, translated_map)
    try:
        out_doc.save(output_path)
        logging.info("Saved DOCX: %s", output_path)
        print(f" DOCX written to: {output_path}")
    except PermissionError:
        from datetime import datetime
        ts_path = (
            Path(output_path).with_stem(
                f"{Path(output_path).stem}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            )
        )
        out_doc.save(str(ts_path))
        logging.info("Saved DOCX (fallback): %s", ts_path)
        print(f" DOCX written to (fallback): {ts_path}")

if __name__ == "__main__":
    run_crew()