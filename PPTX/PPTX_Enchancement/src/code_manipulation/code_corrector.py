import logging, re
from typing import Dict, Any, Optional

from langchain.prompts import PromptTemplate
from config.llmProvider import gemini_flash_llm

CODE_CORRECTION_PROMPT = """
You are an expert AI code correction agent specialized in debugging and fixing JavaScript code snippets for PowerPoint automation using Office.js.

**Goal:**  
Analyze the provided JavaScript code snippet and the error message it produced during execution inside a Docker environment. Identify the root cause of the error and generate a corrected version that achieves the same task without failure.

**Input Context:**

- **Original action and Task Description:**  
  "{action}", "{task_description}" 

- **Failing JavaScript Code Snippet:**  
```javascript
{failing_code}
```
- **Error Message:**  
"{error_message}"

**Instructions:**
1. **Understand the Task:** Review the task description to identify the user's intent.
2. **Analyze the Error Message:** Determine the type and source of the failure.
3. **Inspect the Code:** Examine the failing code snippet to locate potential issues.
4. **Identify the Root Cause:** Clarify if the issue is due to syntax, API misuse, scoping, or logic errors.
5. **Generate Corrected Code:** Fix the issue while preserving the task’s original purpose.
6. **Preserve Original Intent:** Do not alter the high-level goal of the code.
7. **Output:** Return ONLY the corrected JavaScript code, no explanations, markdown formatting, or comments (unless critical).
"""

def generate_code_with_retry(
    failing_code: str,
    error_message: str,
    agent_task_specification: Dict[str, Any]
) -> Optional[str]:
    logging.info(f"Attempting CODE CORRECTION for task: {agent_task_specification.get('action')} with error: {error_message}")
    
    try:
        prompt_template = PromptTemplate.from_template(CODE_CORRECTION_PROMPT)
        prompt_input = {
            "task_description": agent_task_specification.get("task_description"),
            "action": agent_task_specification.get("action"),
            "failing_code": failing_code,
            "error_message": error_message
        }

        chain = prompt_template | gemini_flash_llm
        response = chain.invoke(prompt_input).strip()

        code_block = re.search(r'```javascript\n(.*?)\n```', response, re.DOTALL)
        extracted_code = code_block.group(1) if code_block else response

        logging.info("GENERATED CORRECTED CODE:\n" + extracted_code)
        return extracted_code

    except Exception as e:
        logging.error(f"Code correction failed for action: {agent_task_specification.get('action')}. Reason: {e}", exc_info=True)
        return None
