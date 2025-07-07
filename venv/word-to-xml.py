from crewai import Agent, Task, Crew
from crewai.tools import tool
from llama_parse import LlamaParse
import os
import re
import sys
import xml.etree.ElementTree as ET


# API Keys
os.environ['OPENAI_API_KEY'] = os.environ.get('OPENAI_KEY')
os.environ['LLAMA_CLOUD_API_KEY'] = os.environ.get('LLAMA_CLOUD_API_KEY')

# ---------- Funções como ferramentas da CrewAI ----------

@tool("Converter DOCX para Markdown com LlamaParse")
def convert_docx_to_markdown_llama_parse(path_word_file: str) -> str:
    """
    Converte um arquivo DOCX para Markdown usando LlamaParse.
    """
    markdown_output = LlamaParse(result_type='markdown').load_data(path_word_file)
    return markdown_output[0].text


@tool("Converter Markdown para XML estruturado")
def markdown_to_xml(markdown_text: str, output_file: str) -> str:
    """
    Converte texto Markdown em um arquivo XML com títulos e parágrafos.
    """
    print("--- Debug Info ---", file=sys.stderr)
    print(f"Input text length: {len(markdown_text) if markdown_text else 0}", file=sys.stderr)
    print(f"Output file: {output_file}", file=sys.stderr)

    sections_and_text = re.split(r'(^#+\s.*$)', markdown_text, flags=re.MULTILINE)
    sections_and_text = [s for s in sections_and_text if s is not None and s != '']

    root = ET.Element('document')
    first_title_index = next((i for i, item in enumerate(sections_and_text) if item.strip().startswith('#')), None)

    if first_title_index is None:
        print("Aviso: Nenhum título Markdown encontrado.", file=sys.stderr)
    else:
        for i in range(first_title_index, len(sections_and_text), 2):
            if i + 1 < len(sections_and_text):
                title_raw = sections_and_text[i].strip()
                title = title_raw.lstrip('#').strip()
                raw_text_block = sections_and_text[i + 1]

                section = ET.SubElement(root, 'section')
                title_element = ET.SubElement(section, 'title')
                title_element.text = title

                paragraphs = re.split(r'\n\s*\n+', raw_text_block)
                for para in paragraphs:
                    stripped_para = para.strip()
                    if stripped_para:
                        text_element = ET.SubElement(section, 'text')
                        text_element.text = '\t' + stripped_para
            else:
                print(f"Warning: Título sem texto subsequente: '{sections_and_text[i].strip()}'", file=sys.stderr)

    tree = ET.ElementTree(root)
    if len(root) > 0:
        tree.write(output_file, encoding='utf-8', xml_declaration=True)
        return f"Arquivo XML gerado com sucesso em: {output_file}"
    else:
        return "Aviso: Nenhuma seção criada. XML não foi gerado."

# ---------- Agentes ----------

docx_agent = Agent(
    role="Conversor DOCX para Markdown",
    goal="Converter documentos .docx em Markdown com precisão",
    tools=[convert_docx_to_markdown_llama_parse],
    backstory="Especialista em extração estruturada de texto de documentos Word usando inteligência artificial.",
    verbose=True
)

xml_agent = Agent(
    role="Gerador de XML a partir de Markdown",
    goal="Transformar Markdown em um XML estruturado com seções e parágrafos",
    tools=[markdown_to_xml],
    backstory="Profundo conhecimento em linguagens de marcação e estruturação de documentos digitais.",
    verbose=True
)

# ---------- Execução da Crew ----------

def executar_pipeline(path_docx: str, path_xml_output: str):
    task1 = Task(
        description=f"Converter o arquivo DOCX '{path_docx}' em Markdown.",
        expected_output="Texto Markdown completo extraído do arquivo DOCX.",
        agent=docx_agent
    )

    task2 = Task(
        description=f"Converter o Markdown gerado para XML e salvar no caminho: {path_xml_output}",
        expected_output=f"Arquivo XML salvo em {path_xml_output}.",
        agent=xml_agent
    )

    crew = Crew(
        agents=[docx_agent, xml_agent],
        tasks=[task1, task2],
        verbose=True
    )

    result = crew.kickoff(inputs={"path_word_file": path_docx, "output_file": path_xml_output})
    print(result)

# ---------- Exemplo de execução ----------

if __name__ == "__main__":
    caminho_docx = r"C:\Users\vitor\OneDrive\Desktop\diagramacao\word\teste.docx"
    caminho_xml = r"C:\Users\vitor\OneDrive\Desktop\diagramacao\output\output_text.xml"
    executar_pipeline(caminho_docx, caminho_xml)
