# question_utils.py

import re
import pandas as pd
import time
from io import StringIO

QUESTION_GEN_PROMPT = """
Generate technical multiple-choice questions for the following microskills. 
Each micro-skill should have at least one 1-mark and one 2-mark question. 
Use clear options and mark the correct answer.

Microskills:
{microskills_block}
"""

EXCEL_FORMATTING_PROMPT = """
Format the following questions into an Excel-compatible markdown table. Use these headers exactly:
SerialNo, SectionName, Tag, PositiveMark, NegativeMark, Level, AnswerTime, Instruction, AnswerExplanation, 
Question, QuestionType, CorrectOption, Option1, Option2, Option3, Option4, Option5, 
Criteria1, Percentage1, Criteria2, Percentage2, Criteria3, Percentage3, Criteria4, Percentage4, 
Criteria5, Percentage5.

Rules:
- Use MULTI_CHOICE for MCQs and ESSAY for descriptive.
- Set Level to EASY/MEDIUM/INTENSE based on complexity.
- Leave optional fields blank unless mentioned.
- Use 1 = Option1, 2 = Option2, etc., for CorrectOption.

Questions:
{raw_questions}
"""

def generate_question_paper(microskills_text, query_openai, api_key):
    try:
        # Step 1: Format microskills block
        microskills_block = ""
        for line in microskills_text.strip().split('\n'):
            if '|' in line:
                title, detail = line.strip().split('|', 1)
                microskills_block += f"{title.strip()}: {detail.strip()}\n"

        # Step 2: Generate raw questions
        question_prompt = QUESTION_GEN_PROMPT.format(microskills_block=microskills_block)
        raw_questions, error1 = query_openai(question_prompt, api_key)
        if error1:
            return None, f"Error generating questions: {error1}"

        # Step 3: Format questions to Excel table
        question_list = re.split(r"\n(?=Q\d+\.)", raw_questions.strip())
        halves = (
            [question_list[:len(question_list)//2], question_list[len(question_list)//2:]]
            if len(question_list) > 20 else [question_list]
        )

        dataframes = []

        for block in halves:
            raw_block = '\n'.join(block).strip()
            format_prompt = EXCEL_FORMATTING_PROMPT.format(raw_questions=raw_block)
            markdown_table, error2 = query_openai(format_prompt, api_key)
            if error2:
                return None, f"Error formatting table: {error2}"

            # Parse markdown table to DataFrame
            md_lines = [line for line in markdown_table.splitlines() if line.strip().startswith("|")]
            md_lines = [line for line in md_lines if not set(line.replace("|", "").strip()) <= {"-", " "}]

            clean_md = "\n".join(md_lines)
            df = pd.read_csv(StringIO(clean_md), sep="|", engine="python", skipinitialspace=True)
            df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
            for col in ["SerialNo", "CorrectOption", "PositiveMark", "NegativeMark"]:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors="coerce")

            dataframes.append(df)

        final_df = pd.concat(dataframes, ignore_index=True)
        return final_df, None

    except Exception as e:
        return None, str(e)
