Attribute VB_Name = "Módulo1"
Sub estrutura()
'para declarar variavel no VBA usamos o comando DIM
Dim produto As String
Dim preço As Double
Dim desconto As Double
Dim precofinal As Double

' vamos ultilizar a caixa e entrada inputbox para as variaves
produto = InputBox("Digite o nome do produto", "Produto")
preco = InputBox("Digite o preco do Produto", "Preco")
desconto = InputBox("Digite o desconto do Produto", "Desconto")

precofinal = preco - preco * desconto

Range("A1").Value = produto
Range("A2").Value = preco
Range("A3").Value = desconto
Range("A4").Value = precofinal


End Sub
