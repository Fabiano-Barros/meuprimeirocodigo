import openpyxl
import pyperclip
import pyautogui
from time import sleep

# entrar na planilha
workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']
# Copiar informação de um campo e colar no seu campo correspondente
for linha in sheet_produtos.iter_rows(min_row=2):
    # Nome do Produto
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(129,192,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    # Descrição
    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(130,298,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    # Categoria
    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(124,412,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    # Codigo produto
    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(127,497,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    # Peso
    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(123,583,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    # Dimensões
    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(125,672,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    # Botão
    pyautogui.click(151,717,duration=1)
    sleep(3)
    
    # Preço
    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(124,217,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    # Quantidade em estoque
    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(128,300,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    # Data de validade
    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(129,389,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    # Cor
    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(128,474,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    # Tamanho
    tamanho = linha[10].value
    pyautogui.click(224,556,duration=1)
    if tamanho == 'Pequeno':
        pyautogui.click(163,592,duration=1)
    elif tamanho == 'Médio':
        pyautogui.click(147,612,duration=1)
    else:
        pyautogui.click(152,637,duration=1)
    
    # Material
    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(127,646,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    # Botão proximo
    pyautogui.click(158,708,duration=1)
    
    # Fabricante 
    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(124,232,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    # Pais de origem
    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(126,323,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    # Observações
    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(124,423,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    # Código de Barras
    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(121,541,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    # Localização Armazem
    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(127,626,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    # Botão Concluir
    pyautogui.click(153,686,duration=1)
    
    # Botão de OK
    pyautogui.click(862,172,duration=1)
    
    # Botão Adicionar Mais Um
    pyautogui.click(677,468,duration=1)
    
# Repitir esses passos para outros campos até preencher campos daquela página
# Clicar em próxima
# Repetir os mesmos passos e ir para a próxima página(página 2)
# Repetir os mesmos passos e finalizar o cadastro daquele produto e clicar em concluir
# clicar em ok, para finalizar o processo
# Clicar no ok mais um vez na mensagem de confirmação de salvamento no banco de dados
# Clicar em "adicionar mais um e repetir o processo até finalizar a planilha"
    
    
    
    
    
    
    
    

