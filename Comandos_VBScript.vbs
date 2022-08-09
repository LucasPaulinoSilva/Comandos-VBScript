'***Comandos VBScript focado em automatização de Excel para RPA***'
---------------------------------------------------------------------------------------------------------------------
'Cria uma função para servir como parametro de entrada:
function Inserir nome da função(Inserir nome do parametro)
on error resume next

'Cria um split para separar o diretório de arquivo Excel caso exista mais de um:
Dim Arr
Arr = split (Inserir nome do parametro, "Inserir delimitador ex: ;")

'Abre uma planilha Excel (Em caso de haver mais de uma):
set excelworkbook = excel.workbooks.open(Arr(Adicionar qual o número da planilha que será aberta, ex: 0;1 etc))

'Abre uma planilha Excel única:
set excelworkbook = excel.workbooks.open("Insere diretório da planilha ou parametro")

'Cria um objeto excel:
set excel = createobject("excel.application")

'Executa aplicação Excel em segundo plano:
excel.application.visible = false

'Desabilita janela de informações:
excelworkbook.removePersonalInformation = false

'Seleciona a primeira aba da planilha:
set aba = excelworkbook.worksheets("Inserir nome da aba ou número")
aba.select

'Seleciona o total de linhas da coluna(s) indicada:
ultimalinha = aba.range("A900000").end(-4162).now

'Filtra as células em branco de um determinado indice:
aba.range("A2:A"&ultimalinha).autofilter Inserir número do indice,"="

'Aplica o auto-filtro para células diferentes de branco:
aba.range("A1:A"&ultimalinha).autofilter Inserir número do indice,"<>"

'Copia os dados das células da coluna(s) indicada:
aba.range("A2:A"&ultimalinha).copy

'Cola os dados em uma ou mais colunas:
aba.range("A2:A").pastespecial

'Apaga linhas de uma ou mais colunas:
aba.range("A:A").delete

'Substitui um carcter especifico por outro:
aba.range("A2:A"&ultimalinha).Replace "Inserir o caracter que será substituido","Inserir o caracter que o substituirá"

'Pega o dia, mês ou ano anterior ou posterior:
dataatual = DateAdd("Substitui pelo o que quer d,m,y",Adiciona ou subtrai um periodo ex: -1, +1, etc,,Date())

'Muda os valores de "General" para "Número" de uma coluna(s) ou linha(s) especifica:
aba.range("A2:A"&ultimalinha).NumberFormat = "General"

'Move uma coluna para outra:
excelworkbook.Sheets(1).Columns("A:B").Cut
excelworkbook.Sheets(1).Columns("C:C").Insert -4161 'Movendo colunas A e B para C

'Maximiza tela:
excel.application.WindowState=-4137

'Maximiza tela:
excel.application.WindowState=-4140

'Centraliza texto das células:
aba.range("A2:B"&ultimalinha).HorizontalAlignment = -4108

'Cria nova aba na planilha Excel:
set aba = excelworkbook.worksheets
aba.Add 
aba("Sheet1").Select

'Move aba para o final:
aba("Nome da aba a ser movida").Move , aba(Inserir nome ou número da aba)

'Renomeia aba:
set aba = excelworkbook.worksheets(Inserir nome ou número da aba)

'Aplica o auto-filtro para células em branco:
aba.range("A2:B"&ultimalinha).autofilter 1,"="

'Aplica o auto-filtro para células diferentes de branco:
aba.range("A2:B"&ultimalinha).autofilter 1,"<>"

'Altera formato das colunas para tipo "Data":
aba.range("A2:B"&ultimalinha).NumberFormat = "m/d/yyyy"

'Altera formato das colunas para tipo "Texto":
aba.range("A2:B"&ultimalinha).NumberFormat = "@"

'Altera formato das colunas para tipo "Número":
aba.range("A2:B"&ultimalinha).NumberFormat = "0"

'Altera formato das colunas para tipo "Accouting":
aba.range("A2:B"&ultimalinha).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

'Cria uma nova coluna:
aba.Columns("A").Insert
aba.range("A1").Value = "Inserir nome da coluna"

'Converte xlsx em csv delimitado por vírgula:
excelworkbook.SaveAs Replace(excelworkbook.FullName), ".xlsx", ".csv", 6

'Converte xls em xlsx:
excelworkbook.SaveAs Replace(excelworkbook.FullName), ".xls", ".xlsx", 51

'Converte data para uma string:
Inserir nome da variavel = CDate(Inseir Data)

'Faz o cálculo de range de meses entre uma data e otra:
meses = datediff("m", "dataArquivo,dataAtual")

'Faz algo enquanto a condição existir:
Do While Inseir uma condição. Ex: IsEmpty (Inserir um objeto criado anteriormente. Ex: currentCell)
   Inserir onde será realizada as ações. Ex: aba.range("A2:Z2").Delete 
   Inserir novamente o objeto criado para continuar o loop. Ex: set currentCell = aba.range("B2")
Loop

'Insere uma fórmula na planilha:
aba.range("Inserir a célula onde será alocada a fórmula").FormulaLocal = "Inserir a fórmula. Ex: =CONCATENATE(D2;H2;I2)"

'Realiza o FillDown nas colunas:
aba.range("A2:B"$ultimalinha).FillDown

'Insere delay de 1 seg:
WScript.Sleep 1000

'Salva o arquivo Excel:
excelworkbook.save 
				
'Fecha arquivo excel
excel.quit

'Fecha planilha (aba) Excel:
excelworkbook.close true or false 'para fechar salvando ou não a planilha

'Verifica se o dia escolhi é útil (Ex: Quinto dia):
Dim s
Dim j
Dim D

D = cDate("05" & "/" & Month(Now()) & "/" & Year(Now()))

    For x = 1 To 10
        j = x + 4 & "/" & Month(Now()) & "/" & Year(Now())
        s = Weekday(j)
		
            If s <> 1 then				
                Exit For
            End If
    Next
	
D = CDate(j)

if Weekday(D) = 7 Then
	D = D+1
end if
if Weekday(D) <> 6 Then
	D = D+1
end if
				
'Abre caixa de mensagem
msgbox D

'Verifica a existência de erro e retorna:
If Erro.Number <> 0 Then
    Dim res
    res = "ERRO, número do erro:" & CStr(Erro.Number) & ", Descrição do erro:" & CStr(Erro.Description)
    Inserir nome da função = res
Else
    Inserir nome da função = "1"
End If
end function

