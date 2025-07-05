import scribus
import math
import traceback
import xml.etree.ElementTree as ET
import os
import sys 


def mm_to_pt(mm):
    """Converte milímetros para pontos (1 pt = 1/72 polegada, 1 polegada = 25.4 mm)"""
    return mm * 72.0 / 25.4

def create_base_frame(x, y, width, height, name):
    """
    Cria um frame de texto básico no Scribus.

    Args:
        x (float): Coordenada X da posição superior esquerda do frame.
        y (float): Coordenada Y da posição superior esquerda do frame.
        width (float): Largura do frame.
        height (float): Altura do frame.
        name (str): Nome único para o frame.

    Returns:
        ScribusObject or None: O objeto frame criado se a operação for bem-sucedida,
                               ou None em caso de erro.
    """
    # Verifica se já existe um objeto com o mesmo nome
    if name in scribus.getAllObjects():
        try:
            # Tenta deletar o objeto existente com o mesmo nome
            scribus.deleteObject(name)

        except scribus.ScribusException:
             # Captura exceções específicas do Scribus durante a deleção
             print(f"AVISO: Nao foi possivel deletar objeto existente {name} (ScribusException).")

        except Exception:
             # Captura qualquer outra exceção durante a deleção
             print(f"AVISO: Nao foi possivel deletar objeto existente {name} (General Error).")

    try:
        # Garante altura e largura mínimas para a criação
        min_valid_dim = 0.1 # Pontos (unidade de medida padrão no Scribus)
        if width < min_valid_dim: width = min_valid_dim
        if height < min_valid_dim: height = min_valid_dim
        # Cria o frame de texto com as dimensões e nome especificados
        frame = scribus.createText(x, y, width, height, name)
        return frame

    except scribus.ScribusException as e:
        # Captura exceções específicas do Scribus durante a criação
        print(f"ERRO: Falha ao criar frame {name} em ({x:.2f}, {y:.2f}) ({width:.2f}, {height:.2f}): {e}")
        return None

    except Exception as e:
         # Captura qualquer outra exceção durante a criação
         print(f"ERRO: Erro geral ao criar frame {name} em ({x:.2f}, {y:.2f}) ({width:.2f}, {height:.2f}): {e}")
         return None

def format_text_frame(name, font_size, line_spacing):
    """
    Formata um frame de texto básico no Scribus.

    Args:
        name (str): O nome único do frame de texto a ser formatado.
        font_size (float): O tamanho da fonte a ser aplicado ao texto no frame.
        line_spacing (float): O valor do espaçamento entre linhas a ser aplicado
                              (no modo fixo).

    Returns:
        bool: True se a formatação foi aplicada com sucesso ao frame.
              False se o frame não foi encontrado ou se ocorreu um erro
              durante a formatação.
    """
    # Verifica se o nome é válido e se o objeto existe
    if not name or name not in scribus.getAllObjects():
        # Retorna False se o nome for inválido ou o objeto não existir
        return False

    try:
        # Seleciona o objeto para garantir que as operações se apliquem a ele (embora algumas funções aceitem o nome diretamente)
        scribus.selectObject(name)
        # Define o tamanho da fonte para o objeto especificado
        scribus.setFontSize(font_size, name)
        # Define o modo de espaçamento entre linhas para Fixo (0) para o objeto especificado
        scribus.setLineSpacingMode(0, name) # 0: Espaçamento Fixo entre Linhas
        # Define o valor do espaçamento entre linhas para o objeto especificado
        scribus.setLineSpacing(line_spacing, name)
        # Desseleciona todos os objetos para limpar o estado de seleção
        scribus.deselectAll()
        # Retorna True indicando sucesso
        return True

    except scribus.ScribusException as e:
        # Em caso de erro específico do Scribus, desseleciona tudo, imprime o erro e retorna False
        scribus.deselectAll()
        print(f"ERRO: Falha ao formatar frame {name}: {e}")
        return False

    except Exception as e:
        # Em caso de qualquer outro erro, desseleciona tudo, imprime o erro e retorna False
        scribus.deselectAll()
        print(f"ERRO: Erro geral ao formatar frame {name}: {e}")
        return False

def set_text_and_layout(name, text):
    """
    Define o texto e realiza o layout em um frame nomeado.

    Retorna True se o texto extravasar ou se ocorrer um erro durante o processo.
    Retorna False se o frame especificado não existir.
    """
    if not name or name not in scribus.getAllObjects():
        return False # Não foi possível definir texto ou layout

    try:
        scribus.setText(text, name)
        scribus.layoutText(name) # Força o fluxo de texto e calcula o extravasamento
        overflows = scribus.textOverflows(name, 0)
        return overflows

    except scribus.ScribusException as e:
        print(f"ERRO: Falha ao definir texto/layout para frame {name}: {e}")
        return True # Retorna True em caso de erro, indicando que algo não funcionou como esperado

    except Exception as e:
         print(f"ERRO: Erro geral ao definir texto/layout para frame {name}: {e}")
         return True # Retorna True em caso de erro geral

def get_required_height(name, line_spacing):
    """
    Calcula a altura vertical necessária para que o texto de um frame caiba sem extravasar.

    Considera o espaçamento entre linhas e o preenchimento (padding) do frame.
    Retorna a altura calculada.
    Retorna -1.0 se o frame for inválido, já estiver com extravasamento ou se ocorrer um erro.
    """
    if not name or name not in scribus.getAllObjects():
        return -1.0 # Frame inválido

    try:
        # Garante que o texto esteja disposto para obter métricas precisas
        scribus.layoutText(name)

        if scribus.textOverflows(name, 0):
             return -1.0 # Não é possível calcular se houver extravasamento

        num_linhas = scribus.getTextLines(name)
        if num_linhas < 0: # Não deveria acontecer se não houver extravasamento, mas trata defensivamente
            num_linhas = 0

        distancias = scribus.getTextDistances(name)
        distancia_superior = distancias[2] if len(distancias) > 2 else 0.0
        distancia_inferior = distancias[3] if len(distancias) > 3 else 0.0

        # Obtém o conteúdo real do texto
        text_content = scribus.getText(name)
        # Verdadeiro se text_content não estiver vazio após remover espaços em branco
        has_content_that_needs_space = bool(text_content.strip())

        altura_necessaria = 0.0
        if num_linhas > 0:
             altura_necessaria = (num_linhas * line_spacing) + distancia_superior + distancia_inferior
        elif has_content_that_needs_space:
             # Se nenhuma linha for relatada mas tiver conteúdo, assume que pelo menos o espaço de uma linha é necessário
             # Usa line_spacing para esta estimativa + preenchimento (padding)
             altura_necessaria = line_spacing + distancia_superior + distancia_inferior
        else:
             # Sem conteúdo (ou apenas espaços em branco), apenas o espaço de preenchimento (padding) é tecnicamente necessário
             altura_necessaria = distancia_superior + distancia_inferior


        # Garante uma altura mínima razoável para um frame que tenha conteúdo
        MIN_EFFECTIVE_CONTENT_HEIGHT = line_spacing / 2.0 if line_spacing > 0 else 1.0
        if altura_necessaria < MIN_EFFECTIVE_CONTENT_HEIGHT and has_content_that_needs_space:
             # Se a altura calculada for menor que a mínima para conteúdo, usa a mínima
             altura_necessaria = MIN_EFFECTIVE_CONTENT_HEIGHT
        elif altura_necessaria <= 0 and has_content_that_needs_space:
             # Se a altura calculada for 0 ou menor apesar de ter conteúdo, redefine para a mínima
             altura_necessaria = MIN_EFFECTIVE_CONTENT_HEIGHT

        # Mínimo absoluto para um frame visível (mesmo que vazio ou com apenas espaços)
        if altura_necessaria <= 0:
             altura_necessaria = 1.0

        return altura_necessaria

    except scribus.ScribusException as e:
        print(f"ERRO: Falha ao calcular altura necessaria para frame {name}: {e}")
        return -1.0 # Retorna -1.0 em caso de erro

    except Exception as e:
        print(f"ERRO: Erro geral ao calcular altura necessaria para frame {name}: {e}")
        return -1.0 # Retorna -1.0 em caso de erro geral

def adjust_frame_height(name, line_spacing):
    """
    Ajusta a altura de um frame de texto para acomodar o conteúdo, 
    somente diminuindo e se o frame não apresentar extravasamento inicial.

    Retorna a coordenada Y inferior do frame após o ajuste (ou a original se não ajustado).
    Retorna None se o frame não existir ou se ocorrer um erro grave.
    """
    if not name or name not in scribus.getAllObjects():
        return None # Não é possível ajustar se o frame não existir

    try:
        # Obtém a posição e tamanho atuais antes de um redimensionamento potencial
        try:
             pos = scribus.getPosition(name)
             size = scribus.getSize(name)
             current_bottom_y = pos[1] + size[1]

        except scribus.ScribusException as e:
             print(f"ERRO: Falha ao obter pos/tamanho de {name} para ajuste: {e}")
             return None # Erro crítico, retorna None

        except Exception as e:
             print(f"ERRO: Erro geral ao obter pos/tamanho de {name} para ajuste: {e}")
             return None # Erro crítico, retorna None


        if scribus.textOverflows(name, 0):
            return current_bottom_y # Não é possível ajustar se ainda houver extravasamento

        # Assume que get_required_height está definida em outro lugar e funciona
        required_height = get_required_height(name, line_spacing)

        if required_height < 0: # O cálculo falhou ou a verificação de extravasamento anterior estava errada
             return current_bottom_y # Retorna a parte inferior atual com base no tamanho original

        # Garante que required_height não seja maior que o tamanho atual se não houver extravasamento relatado
        # Apenas diminuímos, não aumentamos além da altura inicial do frame se nenhum extravasamento for detectado.
        if required_height > size[1]:
             required_height = size[1] # Não aumenta o frame se o Scribus disser que não há extravasamento

        # Apenas redimensiona se a altura necessária for significativamente menor
        min_resize_diff = 0.1 # Pontos (unidade padrão do Scribus)
        if 0 < required_height < size[1] - min_resize_diff:

            try:
                scribus.sizeObject(size[0], required_height, name)
                # Obtém a posição novamente após o redimensionamento, pois pode deslocar ligeiramente
                final_pos = scribus.getPosition(name)
                final_size = scribus.getSize(name)
                return final_pos[1] + final_size[1]

            except scribus.ScribusException as e:
                 print(f"ERRO: Falha ao redimensionar {name} para {required_height:.2f}: {e}")
                 return current_bottom_y # O redimensionamento falhou, retorna a parte inferior original

            except Exception as e:
                 print(f"ERRO: Erro geral ao redimensionar {name} para {required_height:.2f}: {e}")
                 return current_bottom_y # O redimensionamento falhou, retorna a parte inferior original
        else:
            # Nenhum ajuste significativo necessário ou a altura necessária é maior (não deveria acontecer sem extravasamento)
            return current_bottom_y # Retorna a parte inferior atual

    except scribus.ScribusException as e:
        print(f"ERRO: Erro durante o processo de ajuste de altura para o frame {name}: {e}")
        return None # Erro crítico, retorna None

    except Exception as e:
        print(f"ERRO: Erro geral durante o processo de ajuste de altura para o frame {name}: {e}")
        return None # Erro crítico, retorna None

def read_xml_file(path_xml):
    """
    Lê o conteúdo de um arquivo XML, buscando por elementos 'section' com seus 'title' e 'text'.

    Retorna uma lista de dicionários, cada um representando uma seção válida encontrada (com título ou texto).
    """
    secoes = []
    try:
        with open(path_xml, 'r', encoding='utf-8') as f:
            tree = ET.parse(f)
        root = tree.getroot()

        for section_xml in root.findall('section'):
            titulo_elem = section_xml.find('title')
            text_elems = section_xml.findall('text')

            titulo_texto = str(titulo_elem.text).strip() if titulo_elem is not None and titulo_elem.text is not None else ""
            # Processa textos, mantendo os vazios se necessário para estrutura, mas pula os que são puramente espaços em branco depois
            list_of_texts = [str(text_elem.text) if text_elem is not None and text_elem.text is not None else "" for text_elem in text_elems]

            # Apenas adiciona a seção se ela tiver um título não vazio ou pelo menos um texto não puramente espaços em branco
            if titulo_texto or any(t.strip() for t in list_of_texts):
                 secoes.append({"titulo": titulo_texto, "textos": list_of_texts})
            else:
                 print(f"DEBUG: Pulando secao vazia (sem titulo e sem texto valido) no XML.")

        print(f"DEBUG: Lido {len(secoes)} secoes validas do XML.")
        return secoes

    except FileNotFoundError:
        scribus.messageBox("Erro", f"Arquivo não encontrado: {path_xml}", icon=scribus.ICON_WARNING)
        print(f"ERRO: Arquivo XML nao encontrado: {path_xml}")
        return []

    except ET.ParseError as e:
        scribus.messageBox("Erro de XML", f"Falha ao analisar o arquivo XML:\n{e}", icon=scribus.ICON_WARNING)
        print(f"ERRO: Falha ao analisar XML: {e}")
        return []

    except Exception as e:
        tb_str = traceback.format_exc()
        scribus.messageBox("Erro Inesperado na Leitura XML", f"{e}\n{tb_str}", icon=scribus.ICON_CRITICAL)
        print(f"ERRO: Erro inesperado na leitura XML: {e}\n{tb_str}")
        return []

def find_bold_font(current_font):
    """
    Tenta encontrar uma variante negrito de uma fonte pelo seu nome, buscando nas fontes disponíveis no Scribus.

    Retorna o nome da fonte negrito encontrada ou None se não for encontrada ou em caso de erro.
    """
    try:
        fontes_disponiveis = scribus.getFontNames()

        nome_base = current_font
        # Remove sufixos comuns para obter o nome base
        for suffix in [" Regular", " Italic", " Bold", " Light", " Thin", " Medium", " Black", " Roman"]:
            if nome_base.endswith(suffix):
                 nome_base = nome_base[:-len(suffix)]
                 break

        # Se já for um tipo de negrito, retorna a fonte atual
        negrito_terms = ["Bold", "Black", "Heavy", "Semibold"]
        if any(term in current_font for term in negrito_terms):
             return current_font

        # Lista de variações negrito comuns baseadas no nome base
        variacoes_negrito_base = [
            nome_base + " Bold",
            nome_base + "-Bold",
            nome_base + " Semibold",
            nome_base + "-Semibold",
            nome_base + " Heavy",
            nome_base + "-Heavy",
            nome_base + " Black",
            nome_base + "-Black",
        ]

        for cand in variacoes_negrito_base:
             if cand in fontes_disponiveis:
                 return cand

        # Tenta adicionar " Bold" ao nome original sem sufixos comuns
        fonte_sem_sufixo = current_font
        for suffix in [" Regular", " Italic", " Roman"]:
             if fonte_sem_sufixo.endswith(suffix):
                  fonte_sem_sufixo = fonte_sem_sufixo[:-len(suffix)]
                  break

        variacao_direta = fonte_sem_sufixo + " Bold"
        if variacao_direta != current_font and variacao_direta in fontes_disponiveis:
             return variacao_direta

        # Busca genérica: fontes contendo o nome base (sem distinção entre maiúsculas/minúsculas)
        # e um dos termos de negrito (sem distinção entre maiúsculas/minúsculas), excluindo a fonte original
        nome_base_lower = nome_base.lower()
        palavras_negrito_lower = ["bold", "semibold", "heavy", "black"]

        for nome_completo in fontes_disponiveis:
             nome_completo_lower = nome_completo.lower()
             # Verifica se o nome base faz parte do nome completo
             if nome_base_lower in nome_completo_lower:
                  # Verifica se algum termo de negrito está no nome completo
                  if any(palavra_negrito in nome_completo_lower for palavra_negrito in palavras_negrito_lower):
                            # Evita retornar a fonte original, a menos que ela já contenha um termo negrito (tratado acima)
                            # Também evita retornar um estilo negrito diferente se a original já era negrito (verificação redundante devido ao primeiro if)
                            if nome_completo != current_font: # Verificação mais simples: apenas não retorna a fonte original
                                 return nome_completo


        # Se nada for encontrado, retorna None
        return None

    except scribus.ScribusException:
        # Em caso de erro específico do Scribus, apenas falha silenciosamente (retornando None)
        return None

    except Exception:
        # Em caso de erro geral, apenas falha silenciosamente (retornando None)
        return None

def get_flattened_items(secoes):
    """
    Achata uma lista estruturada de seções (contendo títulos e textos)
    em uma única lista linear de itens.

    Cada item representa um título ou um bloco de texto e inclui
    seu tipo ('title' ou 'text'), o conteúdo e índices que indicam
    sua seção original e posição relativa dentro dela.
    """
    flattened = []
    for section_index, secao in enumerate(secoes):
        titulo_texto = secao.get("titulo", "").strip()
        # Adiciona item de título apenas se tiver conteúdo
        if titulo_texto:
            flattened.append({
                'type': 'title',
                'text': titulo_texto,
                'section_index': section_index,
                'item_index_in_section': 0 # Títulos são conceitualmente o primeiro item em uma seção
            })

        # Adiciona itens de texto, pulando textos com apenas espaços em branco
        # item_index_in_section conta textos dentro da seção, começando de 0 para o primeiro texto
        text_counter_in_section = 0
        for text_index_xml, text_body in enumerate(secao.get("textos", [])):
             # Adiciona como itens apenas textos que não são só espaços em branco
             if text_body.strip():
                 flattened.append({
                     'type': 'text',
                     'text': text_body,
                     'section_index': section_index,
                     'item_index_in_section': text_counter_in_section # Índice entre os itens de texto *válidos* dentro desta seção
                 })

                 text_counter_in_section += 1

    print(f"DEBUG: Lista final 'flattened_items' tem {len(flattened)} itens.")
    return flattened


def main():
    # Configuração do Documento
    if not scribus.haveDoc():
        largura_pagina_mm = 210
        altura_pagina_mm = 297
        margem_mm = 14.111
        try:
             scribus.newDocument(
                (mm_to_pt(largura_pagina_mm), mm_to_pt(altura_pagina_mm)),
                (mm_to_pt(margem_mm), mm_to_pt(margem_mm), mm_to_pt(margem_mm), mm_to_pt(margem_mm)),
                scribus.PORTRAIT, 1, scribus.UNIT_POINTS,
                scribus.PAGE_1, False, False
            )
             print("DEBUG: Novo documento criado.")
        except Exception as e:
            scribus.messageBox("Erro ao Criar Documento", f"{e}", icon=scribus.ICON_WARNING)
            print(f"ERRO: Falha ao criar documento: {e}")
            return
    else:
        # Limpa o documento existente
        try: scribus.setRedraw(False)
        except: pass
        print("DEBUG: Limpando documento existente.")
        num_pages = scribus.pageCount()
        all_objects = []
        for i in range(1, num_pages + 1):
            try:
                scribus.gotoPage(i)
                all_objects.extend([item[0] for item in scribus.getPageItems()])
            except scribus.ScribusException:
                 print(f"AVISO: Falha ao obter itens da página {i} para limpeza.") 
            except Exception:
                 print(f"AVISO: Erro ao obter itens da página {i} para limpeza.") 

        print(f"DEBUG: Tentando deletar {len(set(all_objects))} objetos.")
        # Usa uma cópia da lista para iterar enquanto deleta
        for obj_name in list(set(all_objects)):
             try:
                 if obj_name in scribus.getAllObjects(): # Verifica se ainda existe
                      scribus.deleteObject(obj_name)
             except scribus.ScribusException:
                 pass
             except Exception:
                 pass

        print("DEBUG: Tentando deletar paginas extras.") 
        # Itera de trás para frente para deletar páginas com segurança
        for p in range(scribus.pageCount(), 1, -1):
             try:
                scribus.gotoPage(p)
                # Verifica se a página está realmente vazia (nenhum item, para simplificar)
                # Uma verificação mais robusta seria necessária se itens de página mestre fossem listados por getPageItems
                items_on_page = scribus.getPageItems()
                if items_on_page:
                    print(f"DEBUG: Pagina {p} contem itens ({len(items_on_page)}). Parando remocao de paginas.") 
                    break
                scribus.deletePage(p)
                print(f"DEBUG: Pagina {p} deletada.") 
             except scribus.ScribusException:
                print(f"AVISO: Nao foi possivel deletar pagina {p} (ScribusException) na finalizacao.") 
                break # Para se não conseguir deletar
             except Exception:
                 print(f"AVISO: Erro geral ao deletar pagina {p} na finalizacao. Parando.") 
                 break # Para se ocorrer erro

        scribus.gotoPage(1)
        try: scribus.setRedraw(True)
        except: pass
        scribus.docChanged(True)
        print("DEBUG: Documento limpo. Pronta para diagramar.") 


    # Seleciona o arquivo XML
    xml_file_path = scribus.fileDialog("Selecione o arquivo XML para diagramar em duas colunas", "*.xml")
    if not xml_file_path:
        scribus.messageBox("Cancelado", "Nenhum arquivo XML selecionado.", icon=scribus.ICON_INFORMATION)
        print("DEBUG: Selecao de arquivo XML cancelada.") 
        return

    # Lê o conteúdo do XML
    secoes_de_conteudo = read_xml_file(xml_file_path)
    if not secoes_de_conteudo:
        scribus.messageBox("Aviso", "Nenhuma seção de conteúdo válida encontrada no arquivo XML.", icon=scribus.ICON_INFORMATION)
        print("DEBUG: Nenhuma secao de conteudo valida encontrada.") 
        return

    # Configurações de Layout
    page_size = scribus.getPageSize()
    altura_pagina_pt = page_size[1]
    margins = scribus.getPageMargins()
    margem_superior_pt = margins[0]
    margem_esquerda_pt = margins[1]
    margem_inferior_pt = margins[2]
    margem_direita_pt = margins[3]

    largura_area_conteudo = page_size[0] - margem_esquerda_pt - margem_direita_pt

    distancia_entre_colunas_mm = 10 # Espaço entre colunas
    distancia_entre_colunas_pt = mm_to_pt(distancia_entre_colunas_mm)

    min_largura_coluna_pt = 10.0 # Largura mínima desejável para uma coluna
    min_largura_para_duas_col = (2 * min_largura_coluna_pt) + distancia_entre_colunas_pt
    if largura_area_conteudo < min_largura_para_duas_col:
         msg = f"Margens horizontais e/ou espaço entre colunas muito grandes para duas colunas.\nLargura da área de conteúdo ({largura_area_conteudo:.2f} pt) insuficiente (requer pelo menos {min_largura_para_duas_col:.2f} pt)."
         scribus.messageBox("Erro de Layout", msg, icon=scribus.ICON_CRITICAL)
         print(f"ERRO: {msg}")
         return

    largura_coluna = (largura_area_conteudo - distancia_entre_colunas_pt) / 2.0

    if largura_coluna <= 1.0: # Verificação de sanidade após o cálculo
         msg = f"A largura calculada da coluna ({largura_coluna:.2f} pt) é muito pequena. Verifique margens e espaco entre colunas."
         scribus.messageBox("Erro de Layout", msg, icon=scribus.ICON_CRITICAL)
         print(f"ERRO: {msg}")
         return

    x_col1 = margem_esquerda_pt
    x_col2 = margem_esquerda_pt + largura_coluna + distancia_entre_colunas_pt

    print(f"DEBUG: Largura da coluna: {largura_coluna:.2f} pt")
    print(f"DEBUG: X Coluna 1: {x_col1:.2f} pt, X Coluna 2: {x_col2:.2f} pt")


    # Configurações de Fonte e Espaçamento
    tamanho_fonte_titulo = 16.0
    espacamento_linha_titulo_fixo = tamanho_fonte_titulo * 1.2
    tamanho_fonte_principal = 12.0
    espacamento_linha_principal_fixo = tamanho_fonte_principal * 1.4

    espaco_vertical_entre_elementos_mm = 4 # Espaço entre itens (título ou texto) dentro da mesma coluna
    espaco_vertical_pt = mm_to_pt(espaco_vertical_entre_elementos_mm)

    espaco_entre_secoes_mm = 8 # Espaço extra antes do primeiro caixa de uma nova seção
    espaco_entre_secoes_pt = mm_to_pt(espaco_entre_secoes_mm)

    # Altura mínima que um frame precisa ter para ser considerado "colocável" em um espaço
    # Usando a altura de uma linha de texto principal + algum preenchimento como referência.
    MIN_PLACEABLE_HEIGHT = espacamento_linha_principal_fixo * 1.5 # Altura de uma linha mais meio espaçamento de linha como mínimo
    if MIN_PLACEABLE_HEIGHT < 5.0: MIN_PLACEABLE_HEIGHT = 5.0 # Mínimo absoluto para visibilidade/clicabilidade

    print(f"DEBUG: MIN_PLACEABLE_HEIGHT (usado): {MIN_PLACEABLE_HEIGHT:.2f} pt")

    # Variáveis de controle de layout
    pagina_atual = 1
    # y_colX_bottom: Posição Y inferior da última caixa colocada em cada coluna na página atual.
    # É o ponto a partir do qual o próximo item naquela coluna tentará iniciar (+ espaçamento).
    y_col1_bottom = margem_superior_pt
    y_col2_bottom = margem_superior_pt

    # Achata as seções de conteúdo em uma única lista ordenada de itens
    flattened_items = get_flattened_items(secoes_de_conteudo)

    last_section_index = -1 # Rastreia o índice da seção do item processado anteriormente


    try:
        scribus.setRedraw(False)

        for item_index, item in enumerate(flattened_items):
            print(f"\n--- Processando Item {item_index + 1}/{len(flattened_items)} (Seção {item.get('section_index', -1)+1}, Tipo: {item.get('type', 'unknown')}) ---")
            scribus.gotoPage(pagina_atual) # Garante que estamos na página correta

            # --- Calcula espaço antes deste item ---
            space_before_this_item = espaco_vertical_pt # Espaço padrão entre itens

            # Verifica se este item inicia uma nova seção (a menos que seja o primeiríssimo item geral)
            is_first_item_overall = (item_index == 0)
            starts_new_section = not is_first_item_overall and item.get('section_index', -1) != last_section_index

            if starts_new_section:
                # Este é o primeiro item de uma nova seção. Adiciona espaço de seção.
                # Este espaço é adicionado *antes* do primeiro frame da nova seção,
                # na página/coluna onde o novo item será realmente colocado.
                # É adicionado ao cálculo de y_start abaixo.
                space_before_this_item = espaco_entre_secoes_pt + espaco_vertical_pt # Espaço de seção + espaço normal entre itens
                print(f"DEBUG: Item inicia nova secao {item.get('section_index', -1)+1}. Adicionando espaco entre secoes ({espaco_entre_secoes_pt:.2f}) + espaco vertical ({espaco_vertical_pt:.2f}). Total: {space_before_this_item:.2f}")
            elif is_first_item_overall:
                 # Este é o primeiríssimo item de todo o documento
                 space_before_this_item = 0.0 # Começa direto na margem superior (nenhum espaço adicionado antes do primeiro item)
                 print(f"DEBUG: Primeiro item geral. Sem espaco antes.")
            else:
                # Mesma seção que o item anterior
                 space_before_this_item = espaco_vertical_pt
                 # print(f"DEBUG: Mesma secao. Espaco vertical padrão ({space_before_this_item:.2f}).")


            last_section_index = item.get('section_index', -1) # Atualiza para o próximo item


            # --- Determina a posição para o PRIMEIRO frame (F1) deste item ---
            # Lógica: Tenta Col 1. Se não houver espaço suficiente, tenta Col 2. Se não houver espaço, Nova Página Col 1.

            # Calcula Y inicial potencial se colocado na Col 1 ou Col 2
            y_start_potential_col1 = y_col1_bottom + space_before_this_item
            y_start_potential_col2 = y_col2_bottom + space_before_this_item

            # Calcula o espaço disponível em cada coluna a partir do Y inicial potencial até a margem inferior
            avail_col1 = altura_pagina_pt - y_start_potential_col1 - margem_inferior_pt
            avail_col2 = altura_pagina_pt - y_start_potential_col2 - margem_inferior_pt

            needs_new_page = False
            col_to_place_f1 = 0 # Coluna onde F1 será realmente colocado
            x_to_place_f1 = 0.0
            y_to_place_f1 = 0.0 # Esta será a coordenada Y final para a criação do frame

            # Verifica se o item cabe minimamente na Col 1
            if avail_col1 >= MIN_PLACEABLE_HEIGHT:
                # Cabe minimamente na Col 1
                col_to_place_f1 = 1
                x_to_place_f1 = x_col1
                y_to_place_f1 = y_start_potential_col1
                print(f"DEBUG: F1 cabe em Col 1 (avail={avail_col1:.2f}). Colocando lá em Y={y_to_place_f1:.2f}.")

            # Se não coube na Col 1, verifica se cabe minimamente na Col 2
            elif avail_col2 >= MIN_PLACEABLE_HEIGHT:
                # Não cabe minimamente na Col 1, mas cabe minimamente na Col 2
                col_to_place_f1 = 2
                x_to_place_f1 = x_col2
                y_to_place_f1 = y_start_potential_col2
                print(f"DEBUG: F1 nao cabe em Col 1 (avail={avail_col1:.2f}), mas cabe em Col 2 (avail={avail_col2:.2f}). Colocando lá em Y={y_to_place_f1:.2f}.")

            else:
                # Não cabe minimamente na página atual de forma alguma. Precisa de nova página.
                needs_new_page = True
                print(f"DEBUG: F1 nao cabe em nenhuma coluna da pag {pagina_atual}. Precisa de nova pagina.")


            # Lida com Nova Página para F1
            if needs_new_page:
                scribus.newPage(-1)
                pagina_atual += 1
                scribus.gotoPage(pagina_atual)
                y_col1_bottom = margem_superior_pt # Reseta os bottoms para a nova página
                y_col2_bottom = margem_superior_pt
                col_to_place_f1 = 1 # Sempre começa na Col 1 em uma nova página
                x_to_place_f1 = x_col1
                y_to_place_f1 = margem_superior_pt # Começa na margem superior na nova página (space_before_this_item é ignorado no topo da nova página)
                print(f"DEBUG: Criada nova pagina {pagina_atual} para F1 de item {item_index+1}. Começando em Y={y_to_place_f1:.2f} na Col {col_to_place_f1}.")


            # Calcula a altura inicial para F1 (altura total restante da coluna a partir de y_to_place_f1)
            altura_para_criar_f1 = altura_pagina_pt - y_to_place_f1 - margem_inferior_pt
            # Garante altura mínima para criação se houver espaço total suficiente na página, senão usa o mínimo absoluto
            # Se a altura restante for minúscula, mas a altura total da coluna for suficiente, cria com altura mínima para permitir o fluxo
            if altura_para_criar_f1 < MIN_PLACEABLE_HEIGHT and (altura_pagina_pt - margem_inferior_pt - margem_superior_pt) >= MIN_PLACEABLE_HEIGHT:
                 altura_para_criar_f1 = MIN_PLACEABLE_HEIGHT
                 print(f"DEBUG: Altura para criar F1 ({altura_para_criar_f1:.2f}) ajustada para MIN_PLACEABLE_HEIGHT.")
            elif altura_para_criar_f1 <= 0:
                 altura_para_criar_f1 = 1.0 # Mínimo absoluto se o espaço for menor que 0
                 print(f"DEBUG: Altura para criar F1 ({altura_para_criar_f1:.2f}) ajustada para 1.0.")


            # --- Cria F1 ---
            # Base de nome mais curta para clareza e menor risco de atingir limites de nome do Scribus
            # Usando abreviação do tipo ('t' para title, 'x' para text)
            item_type_abbr = 't' if item.get('type') == 'title' else 'x'
            item_name_base = f"sec{item.get('section_index', -1)+1}_{item_type_abbr}{item.get('item_index_in_section', -1)+1}"
            f1_name = f"{item_name_base}_p{pagina_atual}_f1_c{col_to_place_f1}"
            print(f"DEBUG: Criando F1 '{f1_name}' em ({x_to_place_f1:.2f}, {y_to_place_f1:.2f}) [{largura_coluna:.2f}x{altura_para_criar_f1:.2f}]")

            f1 = create_base_frame(x=x_to_place_f1, y=y_to_place_f1, width=largura_coluna, height=altura_para_criar_f1, name=f1_name)

            if not f1:
                 print(f"ERRO: Nao foi possivel criar F1 para '{item_name_base}'. Pulando item.")
                 continue # Pula para o próximo item se a criação de F1 falhou


            # Formata F1 e define o texto
            font_size_f1 = tamanho_fonte_principal
            line_spacing_f1 = espacamento_linha_principal_fixo
            if item.get('type') == 'title':
                font_size_f1 = tamanho_fonte_titulo
                line_spacing_f1 = espacamento_linha_titulo_fixo

            format_text_frame(f1, font_size_f1, line_spacing_f1)

            text_to_set_f1 = item.get('text', '')
            # Adiciona recuo de tabulação SOMENTE se o item for do tipo texto E tiver conteúdo após remover espaços em branco
            if item.get('type') == 'text' and text_to_set_f1.strip():
                 text_to_set_f1 = "\t" + text_to_set_f1

            overflows_initial = set_text_and_layout(f1, text_to_set_f1)
            print(f"DEBUG: F1 '{f1_name}' criado e textado. Overflow: {overflows_initial}")

            # Aplica fonte negrito ao título *após* definir o texto
            if item.get('type') == 'title':
                try:
                    current_font_full = scribus.getFont(f1)
                    # Remove possíveis estilos como "Regular", "Normal", etc. para encontrar a família base
                    font_family_base = current_font_full.replace(" Regular", "").replace(" Normal", "").strip()
                    bold_font_name = find_bold_font(font_family_base)
                    if bold_font_name:
                         scribus.selectObject(f1)
                         scribus.setFont(bold_font_name, f1)
                         scribus.deselectAll()
                    else:
                         print(f"AVISO: Nao encontrou fonte negrito para '{font_family_base}'. Titulo '{f1_name}' nao formatado em negrito.")


                except scribus.ScribusException as e:
                     scribus.deselectAll()
                     print(f"AVISO: Falha ao definir fonte negrito para o titulo {f1_name}: {e}")

                except Exception as e:
                     scribus.deselectAll()
                     print(f"AVISO: Erro geral ao definir fonte negrito para o titulo {f1_name}: {e}")


            # Atualiza y_colX_bottom com a posição Y inferior inicial de F1.
            # Isso reserva o espaço que F1 inicialmente ocupa em sua coluna.
            # Esta atualização deve acontecer mesmo se houver overflow, pois define o ponto de início
            # para o próximo frame na cadeia ou o próximo item.
            # current_chain_col não é estritamente necessário fora do loop de vinculação agora,
            # pois a decisão para o F1 do PRÓXIMO item é feita com base no espaço disponível na Col1 e depois na Col2.
            # Mas mantido localmente para debug no loop de vinculação. (Comentário original ajustado)
            try:
                f1_pos = scribus.getPosition(f1)
                f1_size = scribus.getSize(f1)
                if col_to_place_f1 == 1: y_col1_bottom = f1_pos[1] + f1_size[1]
                else: y_col2_bottom = f1_pos[1] + f1_size[1]
                print(f"DEBUG: F1 '{f1_name}' criado. Bottoms inicializados: Col1={y_col1_bottom:.2f}, Col2={y_col2_bottom:.2f}")
            except scribus.ScribusException as e:
                print(f"ERRO: Falha ao obter pos/size inicial para F1 '{f1_name}': {e}. Nao posso inicializar y_col_bottom precisamente.")
                # Fallback: Estima o bottom com base nos valores de criação. Menos confiável.
                if col_to_place_f1 == 1: y_col1_bottom = y_to_place_f1 + altura_para_criar_f1
                else: y_col2_bottom = y_to_place_f1 + altura_para_criar_f1
            except Exception as e:
                print(f"ERRO: Erro geral ao obter pos/size inicial para F1 '{f1_name}': {e}. Nao posso inicializar y_col_bottom precisamente.")
                if col_to_place_f1 == 1: y_col1_bottom = y_to_place_f1 + altura_para_criar_f1
                else: y_col2_bottom = y_to_place_f1 + altura_para_criar_f1


            # --- Loop de vinculação para frames de overflow ---
            last_frame_name = f1 # O último frame criado nesta cadeia
            frame_counter = 1
            MAX_FRAMES_PER_ITEM = 100 # Limite de segurança

            # Continua vinculando enquanto o último frame transborda E nenhum limite foi atingido E last_frame_name é válido
            while last_frame_name and last_frame_name in scribus.getAllObjects() and \
                  scribus.textOverflows(last_frame_name, 0) and \
                  frame_counter < MAX_FRAMES_PER_ITEM:

                frame_counter += 1
                print(f"DEBUG: Item '{item_name_base}' overflows from '{last_frame_name}'. Attempting to create frame {frame_counter}.")
                scribus.gotoPage(pagina_atual) # Garante que estamos na página onde o *último* frame foi criado

                # Determina a coluna do frame que está transbordando e sua posição Y inferior
                last_frame_overflow_col = 0
                last_frame_bottom_y_overflow = None # Será calculado se for bem-sucedido
                try:
                    last_frame_pos_overflow = scribus.getPosition(last_frame_name)
                    last_frame_size_overflow = scribus.getSize(last_frame_name)
                    last_frame_bottom_y_overflow = last_frame_pos_overflow[1] + last_frame_size_overflow[1]
                    # Usa uma pequena tolerância para determinar a coluna
                    if abs(last_frame_pos_overflow[0] - x_col1) < abs(last_frame_pos_overflow[0] - x_col2):
                         last_frame_overflow_col = 1
                    else:
                         last_frame_overflow_col = 2
                    print(f"DEBUG: Overflow check from '{last_frame_name}' (Col {last_frame_overflow_col}, Bottom Y={last_frame_bottom_y_overflow:.2f}).")

                except scribus.ScribusException as e:
                    print(f"ERRO: Falha ao obter pos/size para {last_frame_name} em loop de overflow: {e}. Quebrando cadeia.") # Quebra o loop em caso de erro
                    last_frame_name = None # Quebra o loop em caso de erro
                    break # Sai do loop while
                except Exception as e:
                    print(f"ERRO: Erro geral ao obter pos/size para {last_frame_name} em loop de overflow: {e}. Quebrando cadeia.") # Quebra o loop em caso de erro
                    last_frame_name = None # Quebra o loop em caso de erro
                    break # Sai do loop while

                # Se falhou ao determinar a coluna ou o bottom de overflow, algo está errado, quebra a cadeia
                if last_frame_overflow_col == 0 or last_frame_bottom_y_overflow is None:
                     print(f"ERRO: Nao foi possivel determinar a coluna/bottom de overflow para {last_frame_name}. Quebrando cadeia.")
                     last_frame_name = None # Quebra o loop
                     break # Sai do loop while


                # Determina para onde o próximo frame vinculado deve ir com base na ordem de preenchimento
                next_col_for_chain = 0
                next_x_link = 0.0
                next_y_link = 0.0
                needs_new_page_link = False

                # Verifica o espaço disponível *imediatamente após* o último frame em sua *coluna atual*.
                space_below_overflow_frame_in_its_col = altura_pagina_pt - last_frame_bottom_y_overflow - margem_inferior_pt
                print(f"DEBUG: Espaco disponivel abaixo de {last_frame_name} na Col {last_frame_overflow_col}: {space_below_overflow_frame_in_its_col:.2f} pt.")


                if space_below_overflow_frame_in_its_col >= MIN_PLACEABLE_HEIGHT:
                    # Espaço suficiente *imediatamente após* o frame transbordando em sua coluna atual.
                    # Coloca o próximo frame vinculado exatamente lá.
                    next_col_for_chain = last_frame_overflow_col # Permanece na mesma coluna
                    next_x_link = x_col1 if next_col_for_chain == 1 else x_col2
                    next_y_link = last_frame_bottom_y_overflow # Começa imediatamente após o frame anterior na cadeia
                    print(f"DEBUG: Espaco disponivel na Col {last_frame_overflow_col} apos {last_frame_name}. Continuando cadeia na mesma coluna em Y={next_y_link:.2f}")

                else:
                     # A coluna atual na página atual está cheia *abaixo* do frame transbordando.
                     # Precisa mover para o próximo local disponível com base na ordem de preenchimento da página/coluna (Col 1 -> Col 2 -> Nova Página Col 1).
                     print(f"DEBUG: Coluna {last_frame_overflow_col} cheia na pagina {pagina_atual} apos frame {last_frame_name}. Buscando proximo espaco valido para linked frame.")

                     if last_frame_overflow_col == 1:
                         # Coluna 1 está cheia. O próximo local é a Col 2 na mesma página.
                         next_col_for_chain = 2
                         next_x_link = x_col2
                         # Começa no bottom calculado atual da Col 2 (que pode ser menor do que onde a Col 1 terminou)
                         next_y_link = y_col2_bottom # Começa após o que foi colocado por último na Col 2
                         avail_next_col = altura_pagina_pt - next_y_link - margem_inferior_pt
                         print(f"DEBUG: Tentando Col 2 na pag {pagina_atual} starting at Y={next_y_link:.2f} (bottom={y_col2_bottom:.2f}). Avail: {avail_next_col:.2f}")
                         if avail_next_col < MIN_PLACEABLE_HEIGHT:
                             # Col 2 também está cheia ou muito curta a partir de seu bottom atual. Precisa de nova página Col 1.
                             needs_new_page_link = True
                             print(f"DEBUG: Coluna 2 na pagina {pagina_atual} tambem cheia. Precisa de nova pagina.")
                     else: # last_frame_overflow_col == 2
                         # Coluna 2 está cheia. O próximo local é Nova Página Col 1.
                         needs_new_page_link = True
                         next_col_for_chain = 1 # Padrão para nova página
                         print(f"DEBUG: Coluna 2 na pagina {pagina_atual} cheia. Precisa de nova pagina.")

                     # Lida com Nova Página para frame vinculado
                     if needs_new_page_link:
                         scribus.newPage(-1)
                         pagina_atual += 1
                         scribus.gotoPage(pagina_atual)
                         y_col1_bottom = margem_superior_pt # Reseta os bottoms
                         y_col2_bottom = margem_superior_pt
                         next_col_for_chain = 1 # Sempre Col 1 na nova página
                         next_x_link = x_col1
                         next_y_link = margem_superior_pt # Começa na margem superior na nova página (nenhum espaçamento vertical adicionado aqui)
                         print(f"DEBUG: Criada nova pagina {pagina_atual} para overflow. Starting at Y={next_y_link:.2f} in Col {next_col_for_chain}.")

                # Atualiza a coluna atual para o fluxo da cadeia para onde o próximo frame estará
                # Isso não é estritamente necessário para a lógica de decisão, mas ajuda a rastrear a localização da cadeia.
                # current_chain_col = next_col_for_chain # Removido, usado apenas para debug print agora (Comentário original ajustado)

                # Calcula a altura inicial para o frame vinculado (altura total restante da coluna a partir de next_y_link)
                altura_proximo_frame_link = altura_pagina_pt - next_y_link - margem_inferior_pt
                # Garante altura mínima para criação se houver espaço total suficiente na página, senão usa o mínimo absoluto
                if altura_proximo_frame_link < MIN_PLACEABLE_HEIGHT and (altura_pagina_pt - margem_inferior_pt - margem_superior_pt) >= MIN_PLACEABLE_HEIGHT:
                     altura_proximo_frame_link = MIN_PLACEABLE_HEIGHT

                elif altura_proximo_frame_link <= 0:
                     altura_proximo_frame_link = 1.0 # Mínimo absoluto se o espaço for menor que 0


                # --- Cria o novo frame vinculado ---
                nome_frame_novo = f"{item_name_base}_p{pagina_atual}_f{frame_counter}_c{next_col_for_chain}"
                print(f"DEBUG: Criando frame vinculado '{nome_frame_novo}' em ({next_x_link:.2f}, {next_y_link:.2f}) [{largura_coluna:.2f}x{altura_proximo_frame_link:.2f}].")

                frame_novo = create_base_frame(
                    x=next_x_link,
                    y=next_y_link,
                    width=largura_coluna, # Frames vinculados também têm largura de coluna
                    height=altura_proximo_frame_link,
                    name=nome_frame_novo
                )

                if frame_novo:
                     # Formata o novo frame (deve herdar, mas é boa prática)
                     font_size_link = tamanho_fonte_principal
                     line_spacing_link = espacamento_linha_principal_fixo
                     if item.get('type') == 'title': # Títulos que se estendem por múltiplos frames
                          font_size_link = tamanho_fonte_titulo
                          line_spacing_link = espacamento_linha_titulo_fixo
                     format_text_frame(frame_novo, font_size_link, line_spacing_link)

                     # Reaplica negrito para continuações de título (apenas para itens de título)
                     if item.get('type') == 'title':
                         try:
                             current_font_full = scribus.getFont(frame_novo)
                             # Remove possíveis estilos como "Regular", "Normal", etc. para encontrar a família base
                             font_family_base = current_font_full.replace(" Regular", "").replace(" Normal", "").strip()
                             bold_font_name = find_bold_font(font_family_base)
                             if bold_font_name:
                                 scribus.selectObject(frame_novo)
                                 scribus.setFont(bold_font_name, frame_novo)
                                 scribus.deselectAll()
                             else:
                                 print(f"AVISO: Nao encontrou fonte negrito para '{font_family_base}'. Continuacao do titulo '{nome_frame_novo}' nao formatado em negrito.")

                         except scribus.ScribusException as e:
                              scribus.deselectAll()
                              print(f"AVISO: Falha ao definir fonte negrito para frame vinculado {nome_frame_novo}: {e}")
                         except Exception as e:
                              scribus.deselectAll()
                              print(f"AVISO: Erro geral ao definir fonte negrito para frame vinculado {nome_frame_novo}: {e}")


                     # Vincula os frames
                     try:
                         scribus.linkTextFrames(last_frame_name, frame_novo)
                         # print(f"DEBUG: Frames '{last_frame_name}' e '{frame_novo}' vinculados com sucesso.") # Traduzido


                         # Atualiza y_colX_bottom para a coluna onde o NOVO frame foi colocado (usando sua altura inicial)
                         # Isso é crucial pois atualiza o ponto de início para o *próximo* frame na cadeia (se houver)
                         # e potencialmente para o *próximo item* se este for o último frame.
                         try:
                             initial_link_size = scribus.getSize(frame_novo)
                             initial_link_pos = scribus.getPosition(frame_novo)
                             if next_col_for_chain == 1: y_col1_bottom = initial_link_pos[1] + initial_link_size[1]
                             else: y_col2_bottom = initial_link_pos[1] + initial_link_size[1]

                         except scribus.ScribusException as e:
                             print(f"AVISO: Falha ao obter pos/size inicial para {nome_frame_novo} apos criacao/link: {e}. Bottoms podem estar imprecisos.")

                         except Exception as e:
                             print(f"AVISO: Erro geral ao obter pos/size inicial para {nome_frame_novo} apos criacao/link: {e}. Bottoms podem estar imprecisos.")


                         # O novo frame é agora o último na cadeia
                         last_frame_name = frame_novo

                     except scribus.ScribusException as e:
                         print(f"ERRO: Falha ao vincular frames '{last_frame_name}' e '{frame_novo}': {e}. Quebrando cadeia.") # Quebra o loop em caso de erro
                         last_frame_name = None # Quebra o loop em caso de erro
                         break # Sai do loop while

                     except Exception as e:
                          print(f"ERRO: Erro geral ao vincular frames '{last_frame_name}' e '{frame_novo}': {e}. Quebrando cadeia.") # Quebra o loop em caso de erro
                          last_frame_name = None
                          break # Sai do loop while

                else:
                     print(f"ERRO: Nao foi possivel criar frame vinculado '{nome_frame_novo}'. Quebrando cadeia.") # Quebra o loop se a criação do frame falhar
                     last_frame_name = None # Quebra o loop se a criação do frame falhar
                     break # Sai do loop while

            # --- Após loop de vinculação: Ajusta altura do ÚLTIMO frame e atualiza y_col_bottom ---
            # O último frame da cadeia pode ter espaço extra na parte inferior se o texto terminou dentro dele.
            # Precisamos ajustar sua altura e atualizar o y_col_bottom para a coluna em que ele está.
            final_frame_of_item = last_frame_name # last_frame_name é o último frame criado/vinculado
            if final_frame_of_item and final_frame_of_item in scribus.getAllObjects():
                 print(f"DEBUG: Ajustando altura final para cadeia do item '{item_name_base}' no frame '{final_frame_of_item}'.")

                 # Determina a coluna onde o último frame realmente está
                 final_frame_col = 0
                 try:
                     final_frame_pos = scribus.getPosition(final_frame_of_item)
                     final_frame_pos_x = final_frame_pos[0]
                     # Usa uma pequena tolerância para determinar a coluna
                     if abs(final_frame_pos_x - x_col1) < abs(final_frame_pos_x - x_col2):
                         final_frame_col = 1
                     else:
                         final_frame_col = 2

                 except scribus.ScribusException as e:
                      print(f"ERRO: Nao foi possivel obter a posicao final para o frame '{final_frame_of_item}' para determinar a coluna: {e}.")
                      final_frame_col = 0 # Indica falha ao encontrar a coluna

                 except Exception as e:
                      print(f"ERRO: Erro geral ao obter a posicao final para o frame '{final_frame_of_item}' para determinar a coluna: {e}.")
                      final_frame_col = 0 # Indica falha


                 if final_frame_col != 0: # Se sabemos a coluna final
                      # Determina o espaçamento de linha correto para ajuste com base no tipo de item
                      line_spacing_adj = espacamento_linha_principal_fixo
                      if item.get('type') == 'title':
                           line_spacing_adj = espacamento_linha_titulo_fixo

                      # Tenta ajustar a altura e obter o Y inferior
                      # Será calculado se for bem-sucedido, senão obtém o bottom atual (Comentário original ajustado)
                      final_bottom_y_item = adjust_frame_height(final_frame_of_item, line_spacing_adj)

                      # Usa o bottom ajustado se for bem-sucedido, senão obtém o bottom atual
                      item_chain_bottom_y = None
                      if final_bottom_y_item is not None:
                          item_chain_bottom_y = final_bottom_y_item

                      else:
                          # Se adjust_frame_height retornou None (provavelmente devido a falha de pos/size dentro do ajuste),
                          # tenta obter o bottom Y atual novamente aqui como fallback.
                          try:
                              pos_fb = scribus.getPosition(final_frame_of_item)
                              size_fb = scribus.getSize(final_frame_of_item)
                              item_chain_bottom_y = pos_fb[1] + size_fb[1]
                              print(f"AVISO: Ajuste falhou para '{final_frame_of_item}'. Usando bottom atual: {item_chain_bottom_y:.2f}.")
                          except:
                              # Se obter pos/size atual também falhar, usa o y_col_bottom que foi definido
                              # quando o frame foi criado/vinculado pela primeira vez. Este é o fallback menos preciso.
                              if final_frame_col == 1: item_chain_bottom_y = y_col1_bottom
                              else: item_chain_bottom_y = y_col2_bottom
                              print(f"ERRO CRITICO: Nao foi possivel obter bottom Y AJUSTADO nem ATUAL para '{final_frame_of_item}'. Usando Y_col_bottom ({item_chain_bottom_y:.2f}) como fallback.")


                      # Garante que item_chain_bottom_y não seja None antes de usá-lo para atualizar y_col_bottoms
                      # ESTA É PROVAVELMENTE A LINHA ONDE A INDENTAÇÃO ESTAVA ERRADA (Mantido literal pois é um comentário sobre o código)
                      if item_chain_bottom_y is not None:
                          # Atualiza o y_col_bottom para a coluna onde a CADEIA DESTE ITEM TERMINOU
                          if final_frame_col == 1: y_col1_bottom = item_chain_bottom_y
                          else: y_col2_bottom = item_chain_bottom_y

                          print(f"DEBUG: Item '{item_name_base}' terminou. Bottoms finais atualizados: Col1={y_col1_bottom:.2f}, Col2={y_col2_bottom:.2f}")

                          # O próximo item tentará começar com base nestes y_col_bottoms atualizados.
                          # A lógica de decisão para o PRÓXIMO item (Col 1 vs Col 2) está no início do loop principal,
                          # baseada em qual coluna (1 e depois 2) tem espaço. Não é necessário definir explicitamente 'coluna_atual' aqui.
                          # Não deveria acontecer se os fallbacks funcionarem, mas como segurança final
                          pass # A lógica principal já lida com isso


                      else:
                          # Não deveria acontecer se os fallbacks funcionarem, mas como segurança final
                          print(f"ERRO CRITICO: Nao foi possivel determinar bottom Y final para '{item_name_base}' APOS FALLBACKS. Layout subsequente pode estar incorreto.")


                 else:
                      # Falhou ao determinar a coluna final. Não é possível atualizar y_col_bottom precisamente.
                      print(f"AVISO: Nao foi possivel determinar coluna final para '{item_name_base}'. y_col_bottoms nao atualizados precisamente apos ajuste.")



            else:
                 # O último frame da cadeia não existe (por exemplo, deletado devido a erro durante a vinculação)
                 print(f"AVISO: Ultimo frame '{final_frame_of_item}' para '{item_name_base}' nao existe apos processamento da cadeia.")


        # Fim do loop de itens.

    except Exception as e:
        tb_str = traceback.format_exc()
        scribus.messageBox("Erro Durante Diagramação", f"Ocorreu um erro inesperado durante a diagramação:\n{e}\n\n{tb_str}", icon=scribus.ICON_CRITICAL)
        print(f"ERRO CRITICO: Erro inesperado durante diagramacao: {e}\n{tb_str}")

    finally:
        # Finaliza o documento
        try:
            scribus.setRedraw(True)
            if scribus.haveDoc():
                 if scribus.pageCount() > 0:
                     # Remove páginas vazias no final (vai de trás para frente por segurança)
                     print("DEBUG: Verificando paginas vazias no final.")
                     for p in range(scribus.pageCount(), 1, -1):
                         scribus.gotoPage(p)
                         # Verifica se a página contém algum frame criado por este script
                         # Uma verificação mais simples: há ALGUNS itens na página?
                         items_on_page = scribus.getPageItems()
                         if not items_on_page: # Se não houver itens, considera deletar
                              try:
                                   scribus.deletePage(p)
                                   print(f"DEBUG: Deletando página {p} vazia.")
                              except scribus.ScribusException:
                                   print(f"AVISO: Nao foi possivel deletar pagina {p} (ScribusException) na finalizacao.")
                                   break # Para se não conseguir deletar
                              except Exception:
                                   print(f"AVISO: Erro ao deletar pagina {p} na finalizacao.")
                                   break # Para se ocorrer erro
                         else:
                             print(f"DEBUG: Pagina {p} contem itens. Parando remocao de paginas finais.")
                             break # Para de deletar assim que uma página não vazia é encontrada

                     scribus.gotoPage(1)
                 scribus.deselectAll()
                 scribus.docChanged(True)
                 print(f"DEBUG: Processo Concluído. Documento finalizado.")
                 xml_filename = os.path.basename(xml_file_path) if xml_file_path and os.path.exists(xml_file_path) else "Arquivo XML não especificado"
                 print(f"Arquivo XML processado: {xml_filename}")
                 print(f"Total de {len(secoes_de_conteudo)} seção(ões) processada(s).")
                 print(f"Documento final com {scribus.pageCount()} página(s).")
                 scribus.messageBox("Concluído", f"Script finalizado.\nArquivo: {xml_filename}\n{len(secoes_de_conteudo)} seção(ões) processada(s).\nTotal de {scribus.pageCount()} página(s).", icon=scribus.ICON_INFORMATION)
            else:
                 # Se não houver documento ativo, não é possível finalizar
                 scribus.messageBox("Aviso", "Nenhum documento ativo para finalizar.", icon=scribus.ICON_WARNING)
                 print("DEBUG: Nao ha documento ativo para finalizar.")

        except Exception as e_final:
            tb_str_final = traceback.format_exc()
            try:
                 scribus.setRedraw(True)
                 if scribus.haveDoc():
                     scribus.messageBox("Erro na Finalização", f"Ocorreu um erro durante a finalização:\n{e_final}\n{tb_str_final}", icon=scribus.ICON_WARNING)
                     print(f"ERRO: Erro na finalizacao: {e_final}\n{tb_str_final}")
                 else:
                      print(f"ERRO: Erro na finalizacao (sem documento ativo): {e_final}\n{tb_str_final}")
            except: pass


if __name__ == '__main__':
    # Verifica se o script está sendo executado dentro do ambiente Scribus
    if 'scribus' not in locals() and 'scribus' not in globals():
         # Tenta usar Tkinter para uma mensagem de erro gráfica fora do Scribus
         try:
             import Tkinter as tk # Python 2
             from tkMessageBox import showerror
         except ImportError:
             try:
                import tkinter as tk # Python 3
                from tkinter.messagebox import showerror
             except ImportError:
                tk = None # Tkinter não disponível

         if tk:
             root = tk.Tk()
             root.withdraw() # Esconde a janela principal
             showerror("Execution Error", "Este script deve ser executado dentro do ambiente Scribus.")
         else:
             # Fallback para print se Tkinter não estiver disponível
             print("Error: Este script deve ser executado dentro do ambiente Scribus.")
         sys.exit(1) # Sai do script se não estiver no Scribus

    # Verifica se o ambiente Scribus foi inicializado corretamente
    elif not hasattr(scribus, 'newDocument'):
         if 'scribus' in locals() or 'scribus' in globals():
             if hasattr(scribus, 'messageBox'):
                 scribus.messageBox("Initialization Error", "Ambiente Scribus não inicializado corretamente.", icon=scribus.ICON_CRITICAL)
             else:
                  print("Error: Ambiente Scribus não inicializado corretamente.")
         else:
              print("Error: Ambiente Scribus não disponível.") # Não deveria acontecer com base na verificação externa
         sys.exit(1) # Sai do script se a API do Scribus não estiver pronta

    else:
         # O ambiente está pronto, executa a função principal
         try:
             main()
         except scribus.ScribusException as se:
              # Captura exceções específicas da API do Scribus
              tb_str = traceback.format_exc()
              msg = f"Scribus API Error:\n{se}\n\nTraceback:\n{tb_str}"
              print(msg)
              try:
                  if 'scribus' in locals() or 'scribus' in globals():
                       if hasattr(scribus, 'messageBox'):
                            scribus.messageBox("Scribus Error", msg, icon=scribus.ICON_WARNING)
              except: pass # Evita falhas no próprio manipulador de erros
              sys.exit(1) # Sai em caso de Exceção do Scribus
         except Exception as e:
              # Captura outras exceções inesperadas
              tb_str = traceback.format_exc()
              msg = f"Critical Unexpected Error:\n{e}\n\nTraceback:\n{tb_str}"
              print(msg)
              try:
                  if 'scribus' in locals() or 'scribus' in globals():
                       if hasattr(scribus, 'messageBox'):
                           scribus.messageBox("Critical Error", msg, icon=scribus.ICON_CRITICAL)
              except:
                  pass # Evita falhas no próprio manipulador de erros
              sys.exit(1) # Sai em caso de qualquer outra Exceção