import scribus
import math
import traceback
import xml.etree.ElementTree as ET
import os


def mm_to_pt(mm):
    """Converte milímetros para pontos (1 pt = 1/72 polegada, 1 polegada = 25.4 mm)"""
    return mm * 72.0 / 25.4

def create_formatted_frame(
    text, 
    frame_name,
    x, 
    initial_y, 
    width,
    height_to_create,
    font_size, 
    fixed_line_spacing,
    space_before=0.0,
    adjust_height_to_content=False
):
    """
    Cria uma caixa de texto formatada (quadro de texto) no Scribus.

    Args:
        text (str): O texto a ser inserido no quadro.
        frame_name (str): O nome único para identificar o quadro no Scribus.
                        Objetos com o mesmo nome serão deletados antes da criação.
        x (float): Posição horizontal (coordenada X) do canto superior esquerdo
                do quadro.
        initial_y (float): Posição vertical inicial (coordenada Y) do canto
                        superior esquerdo do quadro *antes* de considerar
                        `space_before`.
        width (float): A largura do quadro de texto.
        height_to_create (float): A altura inicial desejada para o quadro. A
                                altura final pode ser diferente se
                                `adjust_height_to_content` for True e não
                                houver estouro.
        font_size (float): O tamanho da fonte para o texto no quadro.
        fixed_line_spacing (float): O espaçamento entre linhas (leading) fixo
                                        a ser aplicado.
        space_before (float, optional): Espaço vertical adicional a ser adicionado
                                        *acima* de `initial_y` para determinar a
                                        posição Y real de criação do quadro.
                                        Padrão é 0.0.
        adjust_height_to_content (bool, optional): Se True, a altura do quadro
                                                    será ajustada para caber o
                                                    texto (mais preenchimento) se
                                                    o texto não estourar e a
                                                    altura calculada for menor que
                                                    `altura_para_criar`. Padrão é
                                                    False.

    Returns:
        Tuple[Optional[str], float, float]: Uma tupla contendo:
            - O nome do quadro criado (`nome_caixa`) em caso de sucesso, ou `None`
            em caso de falha (altura inválida ou erro).
            - A altura *final* do quadro criado em caso de sucesso, ou `0.0` em
            caso de falha.
            - A coordenada Y *depois* do quadro criado
            (y_inicial + espaco_antes + altura_final) em caso de sucesso, ou a
            `y_inicial` original em caso de falha.
            Esta última coordenada Y é útil para posicionar o próximo objeto.
    """
    y_caixa = initial_y + space_before

    MIN_HEIGHT = max(1.0, font_size / 3.0)
    if height_to_create <= MIN_HEIGHT:
        return None, 0.0, initial_y

    try:
        if frame_name in scribus.getAllObjects():
            try:
                 scribus.deleteObject(frame_name)
            except scribus.ScribusException:
                 pass

        caixa = scribus.createText(x, y_caixa, width, height_to_create, frame_name)
        if not caixa:
             return None, 0.0, initial_y

        scribus.selectObject(frame_name)
        scribus.setFontSize(font_size, frame_name)
        scribus.setLineSpacingMode(0, frame_name)
        scribus.setLineSpacing(fixed_line_spacing, frame_name)

        distancias = scribus.getTextDistances(frame_name)
        distancia_superior = distancias[2] if len(distancias) > 2 else 0.0
        distancia_inferior = distancias[3] if len(distancias) > 3 else 0.0

        scribus.setText(text, frame_name)
        scribus.layoutText(frame_name)

        altura_final_caixa = height_to_create

        if adjust_height_to_content and not scribus.textOverflows(frame_name, 0):
            num_linhas = scribus.getTextLines(frame_name)
            if num_linhas < 0: num_linhas = 0

            altura_necessaria = 0.0
            if num_linhas > 0:
                altura_necessaria = (num_linhas * fixed_line_spacing) + distancia_superior + distancia_inferior
            elif text.strip():
                 altura_necessaria = fixed_line_spacing + distancia_superior + distancia_inferior
            else:
                altura_necessaria = distancia_superior + distancia_inferior

            if altura_necessaria <= 0 and height_to_create > MIN_HEIGHT:
                 altura_necessaria = MIN_HEIGHT

            if 0 < altura_necessaria < height_to_create:
                 try:
                     scribus.sizeObject(width, altura_necessaria, frame_name)
                     altura_final_caixa = altura_necessaria
                 except scribus.ScribusException as e_size:
                      altura_final_caixa = height_to_create
                 except Exception as e_size_gen:
                      altura_final_caixa = height_to_create

        y_final_caixa = initial_y + space_before + altura_final_caixa

        scribus.deselectAll()
        return frame_name, altura_final_caixa, y_final_caixa

    except scribus.ScribusException as e:
        if frame_name and frame_name in scribus.getAllObjects():
            try: scribus.deleteObject(frame_name)
            except: pass
        return None, 0.0, initial_y
    except Exception as e:
        tb_str = traceback.format_exc()
        if frame_name and frame_name in scribus.getAllObjects():
            try: scribus.deleteObject(frame_name)
            except: pass
        return None, 0.0, initial_y

def read_xml_file(xml_path):
    """
    Lê o conteúdo de um arquivo XML, buscando por elementos 'section' com seus 'title' e 'text'.

    Retorna uma lista de dicionários, cada um representando uma seção válida encontrada (com título ou texto).
    """
    secoes = []
    try:
        with open(xml_path, 'r', encoding='utf-8') as f:
            tree = ET.parse(f)
        root = tree.getroot()

        for section_xml in root.findall('section'):
            titulo_elem = section_xml.find('title')
            text_elems = section_xml.findall('text')

            titulo_texto = str(titulo_elem.text).strip() if titulo_elem is not None and titulo_elem.text is not None else ""
            list_of_texts = [str(text_elem.text).strip() for text_elem in text_elems if text_elem is not None and text_elem.text is not None]

            secoes.append({"titulo": titulo_texto, "textos": list_of_texts})

        return secoes

    except FileNotFoundError:
        scribus.messageBox("Erro", f"{xml_path}", icon=scribus.ICON_WARNING)
        return []
    except ET.ParseError as e:
        scribus.messageBox("Erro de XML", f"{e}", icon=scribus.ICON_WARNING)
        return []
    except Exception as e:
        tb_str = traceback.format_exc()
        scribus.messageBox("Erro Inesperado", f"{e}", icon=scribus.ICON_CRITICAL)
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


def main():
    if not scribus.haveDoc():
        largura_pagina_mm = 210
        altura_pagina_mm = 297
        largura_pagina_pt = mm_to_pt(largura_pagina_mm)
        altura_pagina_pt = mm_to_pt(altura_pagina_mm)
        margem_mm = 14.111
        margem_pt = mm_to_pt(margem_mm)
        try:
             scribus.newDocument(
                (largura_pagina_pt, altura_pagina_pt),
                (margem_pt, margem_pt, margem_pt, margem_pt),
                scribus.PORTRAIT, 1, scribus.UNIT_POINTS,
                scribus.PAGE_1, False, False
            )
        except Exception as e:
            scribus.messageBox("Erro", f"{e}", icon=scribus.ICON_WARNING)
            return
    else:
        try: scribus.setRedraw(False)
        except: pass

        num_pages = scribus.pageCount()
        all_objects = []
        for i in range(1, num_pages + 1):
            scribus.gotoPage(i)
            all_objects.extend([item[0] for item in scribus.getPageItems()])

        for obj_name in all_objects:
             try:
                 if obj_name in scribus.getAllObjects():
                      scribus.deleteObject(obj_name)
             except scribus.ScribusException:
                 pass

        while scribus.pageCount() > 1:
             try:
                scribus.deletePage(scribus.pageCount())
             except scribus.ScribusException as e:
                break
             except Exception as e:
                 break

        scribus.gotoPage(1)
        try: scribus.setRedraw(True)
        except: pass
        scribus.docChanged(True)

    xml_file_path = scribus.fileDialog("Selecione o arquivo XML", "*.xml")
    if not xml_file_path:
        scribus.messageBox("Cancelado", "Nenhum arquivo XML selecionado.", icon=scribus.ICON_INFORMATION)
        return

    secoes_de_conteudo = read_xml_file(xml_file_path)
    if not secoes_de_conteudo:
        scribus.messageBox("Aviso", "Nenhuma seção de conteúdo válida encontrada no arquivo XML.", icon=scribus.ICON_INFORMATION)
        return

    page_size = scribus.getPageSize()
    largura_pagina_pt = page_size[0]
    altura_pagina_pt = page_size[1]
    margins = scribus.getPageMargins()
    margem_superior_pt = margins[0]
    margem_esquerda_pt = margins[1]
    margem_inferior_pt = margins[2]
    margem_direita_pt = margins[3]

    tamanho_fonte_titulo = 16.0
    espacamento_linha_titulo_fixo = tamanho_fonte_titulo * 1.2
    tamanho_fonte_principal = 12.0
    espacamento_linha_principal_fixo = tamanho_fonte_principal * 1.4
    espaco_vertical_entre_elementos_mm = 5
    espaco_vertical_pt = mm_to_pt(espaco_vertical_entre_elementos_mm)
    espaco_entre_secoes_mm = 10
    espaco_entre_secoes_pt = mm_to_pt(espaco_entre_secoes_mm)

    try:
        scribus.setRedraw(False)
    except: pass

    largura_caixa_comum = largura_pagina_pt - margem_esquerda_pt - margem_direita_pt
    x_caixa = margem_esquerda_pt

    if largura_caixa_comum <= 1.0:
         scribus.messageBox("Erro", "Margens horizontais muito grandes.", icon=scribus.ICON_WARNING)
         try: scribus.setRedraw(True)
         except: pass
         return

    y_cursor = margem_superior_pt
    pagina_atual = 1
    MAX_OVERFLOW_PAGES_PER_TEXT = 50

    for indice_secao, secao in enumerate(secoes_de_conteudo):
        scribus.gotoPage(pagina_atual)

        espaco_antes_desta_secao_pt = 0.0
        if indice_secao > 0:
             if y_cursor > margem_superior_pt + 1.0:
                 espaco_antes_desta_secao_pt = espaco_entre_secoes_pt

        altura_restante_pagina_antes_titulo = altura_pagina_pt - y_cursor - margem_inferior_pt

        titulo_texto = secao.get("titulo", "")
        if titulo_texto.strip():
            min_altura_titulo = espacamento_linha_titulo_fixo + (tamanho_fonte_titulo/3.0) * 2
            if altura_restante_pagina_antes_titulo - espaco_antes_desta_secao_pt < min_altura_titulo:
                 scribus.newPage(-1)
                 pagina_atual += 1
                 scribus.gotoPage(pagina_atual)
                 y_cursor = margem_superior_pt
                 altura_restante_pagina_antes_titulo = altura_pagina_pt - y_cursor - margem_inferior_pt
                 espaco_antes_desta_secao_pt = 0.0

            nome_caixa_titulo = f"sec_{indice_secao+1}_titulo_p{pagina_atual}"

            nome_caixa_titulo_criada, altura_titulo_usada, y_cursor_apos_titulo = create_formatted_frame(
                text=titulo_texto,
                frame_name=nome_caixa_titulo,
                x=x_caixa,
                initial_y=y_cursor,
                width=largura_caixa_comum,
                height_to_create=altura_restante_pagina_antes_titulo,
                font_size=tamanho_fonte_titulo,
                fixed_line_spacing=espacamento_linha_titulo_fixo,
                space_before=espaco_antes_desta_secao_pt,
                adjust_height_to_content=True
            )

            if nome_caixa_titulo_criada:
                 y_cursor = y_cursor_apos_titulo
                 try:
                     if nome_caixa_titulo_criada in scribus.getAllObjects():
                         scribus.selectObject(nome_caixa_titulo_criada)
                         current_font_full = scribus.getFont(nome_caixa_titulo_criada)
                         bold_font_name = find_bold_font(current_font_full)
                         if bold_font_name:
                             scribus.setFont(bold_font_name, nome_caixa_titulo_criada)
                         scribus.deselectAll()
                     else:
                          print(f"{nome_caixa_titulo_criada}'")
                 except scribus.ScribusException as e_font:
                     scribus.deselectAll()
                 except Exception as e_gen_font:
                     tb_str_font = traceback.format_exc()
                     scribus.deselectAll()
            else:
                 print(f"'{nome_caixa_titulo}' não foi criada.")


        y_cursor_apos_item_anterior = y_cursor

        textos_desta_secao = secao.get("textos", [])

        for indice_texto, texto_corpo in enumerate(textos_desta_secao):

            texto_para_diagramar = "\t" + texto_corpo if texto_corpo.strip() else ""

            if not texto_para_diagramar.strip():
                 continue

            espaco_antes_deste_texto_pt = espaco_vertical_pt

            y_pos_inicial_com_espaco = y_cursor_apos_item_anterior + espaco_antes_deste_texto_pt
            altura_restante_pagina_para_esta_caixa = altura_pagina_pt - y_pos_inicial_com_espaco - margem_inferior_pt

            min_altura_texto = espacamento_linha_principal_fixo + (tamanho_fonte_principal/3.0) * 2
            if altura_restante_pagina_para_esta_caixa <= (tamanho_fonte_principal / 3.0):
                 scribus.newPage(-1)
                 pagina_atual += 1
                 scribus.gotoPage(pagina_atual)
                 y_cursor_apos_item_anterior = margem_superior_pt
                 y_pos_inicial_com_espaco = y_cursor_apos_item_anterior
                 altura_restante_pagina_para_esta_caixa = altura_pagina_pt - y_pos_inicial_com_espaco - margem_inferior_pt
                 if altura_restante_pagina_para_esta_caixa <= (tamanho_fonte_principal / 3.0):
                     scribus.messageBox("Erro de Layout", f"Não há espaço suficiente na página {pagina_atual}.", icon=scribus.ICON_CRITICAL)
                     break

            altura_primeiro_frame = altura_restante_pagina_para_esta_caixa
            nome_base_texto = f"sec_{indice_secao+1}_txt_{indice_texto+1}"
            nome_primeiro_frame = f"{nome_base_texto}_p{pagina_atual}_f1"

            y_inicial_real_caixa = y_pos_inicial_com_espaco
            if y_cursor_apos_item_anterior <= margem_superior_pt + 1.0:
                 y_inicial_real_caixa = y_cursor_apos_item_anterior

            primeiro_frame_criado, altura_primeiro_frame_usada, y_pos_apos_primeiro_frame = create_formatted_frame(
                text=texto_para_diagramar,
                frame_name=nome_primeiro_frame,
                x=x_caixa,
                initial_y=y_inicial_real_caixa,
                width=largura_caixa_comum,
                height_to_create=altura_primeiro_frame,
                font_size=tamanho_fonte_principal,
                fixed_line_spacing=espacamento_linha_principal_fixo,
                space_before=0.0,
                adjust_height_to_content=False
            )

            if not primeiro_frame_criado:
                break

            y_cursor_apos_item_anterior = y_pos_apos_primeiro_frame

            frame_anterior_no_fluxo = primeiro_frame_criado
            last_frame_in_chain = primeiro_frame_criado
            contador_frames_vinculados = 1
            paginas_overflow_criadas = 0

            while last_frame_in_chain and last_frame_in_chain in scribus.getAllObjects() and \
                  scribus.textOverflows(last_frame_in_chain, 0) and \
                  paginas_overflow_criadas < MAX_OVERFLOW_PAGES_PER_TEXT:

                paginas_overflow_criadas += 1
                contador_frames_vinculados += 1

                scribus.newPage(-1)
                pagina_atual += 1
                scribus.gotoPage(pagina_atual)

                x_nova_caixa = margem_esquerda_pt
                y_nova_caixa = margem_superior_pt
                altura_nova_caixa = altura_pagina_pt - margem_superior_pt - margem_inferior_pt

                if altura_nova_caixa <= 1.0 or largura_caixa_comum <= 1.0:
                     scribus.messageBox("Erro de Layout", f"Margens da página {pagina_atual} são muito grandes.", icon=scribus.ICON_CRITICAL)
                     last_frame_in_chain = None
                     break

                nome_frame_novo = f"{nome_base_texto}_p{pagina_atual}_f{contador_frames_vinculados}"
                if nome_frame_novo in scribus.getAllObjects():
                     try: scribus.deleteObject(nome_frame_novo)
                     except: pass

                try:
                    frame_novo = scribus.createText(x_nova_caixa, y_nova_caixa, largura_caixa_comum, altura_nova_caixa, nome_frame_novo)
                    if not frame_novo:
                         last_frame_in_chain = None
                         break

                    scribus.selectObject(nome_frame_novo)
                    scribus.setFontSize(tamanho_fonte_principal, nome_frame_novo)
                    scribus.setLineSpacingMode(0, nome_frame_novo)
                    scribus.setLineSpacing(espacamento_linha_principal_fixo, nome_frame_novo)
                    scribus.deselectAll()

                    if frame_anterior_no_fluxo and frame_anterior_no_fluxo in scribus.getAllObjects():
                         try:
                             scribus.linkTextFrames(frame_anterior_no_fluxo, frame_novo)
                         except scribus.ScribusException as e_link:
                             last_frame_in_chain = None
                             break
                    else:
                         last_frame_in_chain = None
                         break

                    frame_anterior_no_fluxo = frame_novo
                    last_frame_in_chain = frame_novo

                except scribus.ScribusException as e_create_link:
                    last_frame_in_chain = None
                    break
                except Exception as e_gen:
                     tb_str_gen = traceback.format_exc()
                     last_frame_in_chain = None
                     break

            if paginas_overflow_criadas >= MAX_OVERFLOW_PAGES_PER_TEXT:
                 scribus.messageBox("Aviso", f"Limite de {MAX_OVERFLOW_PAGES_PER_TEXT}", icon=scribus.ICON_WARNING)

            if last_frame_in_chain and last_frame_in_chain in scribus.getAllObjects():
                 try:
                     scribus.layoutText(last_frame_in_chain)

                     if not scribus.textOverflows(last_frame_in_chain, 0):
                         num_linhas_final = scribus.getTextLines(last_frame_in_chain)

                         if num_linhas_final >= 0:
                             distancias_final = scribus.getTextDistances(last_frame_in_chain)
                             dist_sup_final = distancias_final[2] if len(distancias_final) > 2 else 0.0
                             dist_inf_final = distancias_final[3] if len(distancias_final) > 3 else 0.0

                             altura_necessaria_final = (num_linhas_final * espacamento_linha_principal_fixo) + dist_sup_final + dist_inf_final

                             if altura_necessaria_final <= 0 and texto_para_diagramar.strip():
                                 altura_necessaria_final = espacamento_linha_principal_fixo + dist_sup_final + dist_inf_final

                             if altura_necessaria_final <= 0: altura_necessaria_final = 1.0

                             pos_final = scribus.getPosition(last_frame_in_chain)
                             size_atual_final = scribus.getSize(last_frame_in_chain)

                             if 0 < altura_necessaria_final < size_atual_final[1]:
                                 try:
                                     scribus.sizeObject(size_atual_final[0], altura_necessaria_final, last_frame_in_chain)
                                     y_cursor_apos_item_anterior = pos_final[1] + altura_necessaria_final
                                 except scribus.ScribusException as e_adj_size:
                                     y_cursor_apos_item_anterior = pos_final[1] + size_atual_final[1]
                                 except Exception as e_adj_size_gen:
                                      y_cursor_apos_item_anterior = pos_final[1] + size_atual_final[1]
                             else:
                                 y_cursor_apos_item_anterior = pos_final[1] + size_atual_final[1]
                         else:
                              try:
                                  pos_final = scribus.getPosition(last_frame_in_chain)
                                  size_atual_final = scribus.getSize(last_frame_in_chain)
                                  y_cursor_apos_item_anterior = pos_final[1] + size_atual_final[1]
                              except:
                                  y_cursor_apos_item_anterior = altura_pagina_pt - margem_inferior_pt
                     else:
                         try:
                             pos_final = scribus.getPosition(last_frame_in_chain)
                             size_atual_final = scribus.getSize(last_frame_in_chain)
                             y_cursor_apos_item_anterior = pos_final[1] + size_atual_final[1]
                         except:
                             y_cursor_apos_item_anterior = altura_pagina_pt - margem_inferior_pt
                 except scribus.ScribusException as e_adjust:
                      try:
                          pos_final = scribus.getPosition(last_frame_in_chain)
                          size_atual_final = scribus.getSize(last_frame_in_chain)
                          y_cursor_apos_item_anterior = pos_final[1] + size_atual_final[1]
                      except:
                          y_cursor_apos_item_anterior = altura_pagina_pt - margem_inferior_pt
                 except Exception as e_gen:
                      tb_str_gen = traceback.format_exc()
                      try:
                          pos_final = scribus.getPosition(last_frame_in_chain)
                          size_atual_final = scribus.getSize(last_frame_in_chain)
                          y_cursor_apos_item_anterior = pos_final[1] + size_atual_final[1]
                      except:
                          y_cursor_apos_item_anterior = altura_pagina_pt - margem_inferior_pt
            else:
                 pass

        y_cursor = y_cursor_apos_item_anterior


    try:
        scribus.setRedraw(True)
        if scribus.haveDoc():
             if scribus.pageCount() > 0:
                 scribus.gotoPage(1)
             scribus.deselectAll()
             scribus.docChanged(True)
             print(f"Processo Concluído.")
             print(f"Arquivo XML processado: {os.path.basename(xml_file_path)}")
             print(f"Total de {len(secoes_de_conteudo)} seção(ões) processada(s).")
             print(f"Documento final com {scribus.pageCount()} página(s).")
             scribus.messageBox("Concluído", f"Script finalizado.\nArquivo: {os.path.basename(xml_file_path)}\n{len(secoes_de_conteudo)} seção(ões) processada(s).\nTotal de {scribus.pageCount()} página(s).", icon=scribus.ICON_INFORMATION)
        else:
             scribus.messageBox("Aviso", "Nenhum documento ativo para finalizar.", icon=scribus.ICON_WARNING)

    except Exception as e:
        tb_str = traceback.format_exc()
        try:
             scribus.setRedraw(True)
             if scribus.haveDoc():
                 scribus.messageBox("Erro na Finalização", f"\n{e}.", icon=scribus.ICON_WARNING)
        except: pass


if __name__ == '__main__':
    try:
        if 'scribus' not in locals() and 'scribus' not in globals():
             try:
                 import Tkinter as tk
                 from tkMessageBox import showerror
             except ImportError:
                 try:
                    import tkinter as tk
                    from tkinter.messagebox import showerror
                 except ImportError:
                    tk = None
             if tk:
                 root = tk.Tk()
                 root.withdraw()
                 showerror("Erro", "Execute este script dentro do Scribus.")

        elif not hasattr(scribus, 'newDocument'):
             scribus.messageBox("Erro de Inicialização", "Ambiente Scribus não inicializado corretamente.", icon=scribus.ICON_CRITICAL)
        else:
             main()

    except scribus.ScribusException as se:
         tb_str = traceback.format_exc()
         msg = f"Erro:\n{se}\n\nTraceback:\n{tb_str}"
         print(msg)
         try: scribus.messageBox("Erro Scribus", msg, icon=scribus.ICON_WARNING)
         except: pass
    except Exception as e:
        tb_str = traceback.format_exc()
        msg = f"Erro:\n{e}\n\nTraceback:\n{tb_str}"
        print(msg)
        try:
            if 'scribus' in locals() or 'scribus' in globals():
                if hasattr(scribus, 'messageBox'):
                    scribus.messageBox("Erro Crítico", msg, icon=scribus.ICON_CRITICAL)
        except:
            pass