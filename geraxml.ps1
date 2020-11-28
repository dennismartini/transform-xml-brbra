#A codificação padrão precisa ser UTF8

#Configuração de requisitos
Add-Type -AssemblyName System.Windows.Forms #Adiciona requisito para criar tela de seleção de arquivo
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = [Environment]::GetFolderPath('Desktop') } #Cria função de selecionar arquivos
$null = $FileBrowser.ShowDialog() #Abre o dialogo para selecionar o arquivo e salva o caminho em FileBrowser.FileName

#Cria função que adquire pasta onde o arquivo será salvo
Function Get-Folder($initialDirectory="") {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")|Out-Null

    $foldername = New-Object System.Windows.Forms.FolderBrowserDialog
    $foldername.Description = "Selecione a pasta para salvar o novo arquivo"
    $foldername.rootfolder = "MyComputer"
    $foldername.SelectedPath = $initialDirectory

    if($foldername.ShowDialog() -eq "OK")
    {
        $folder += $foldername.SelectedPath
    }
    return $folder
}

$file = $FileBrowser.FileName #alimenta o caminho do arquivo que será usado para adquirir o arquivo em $xmldoc
$folder = Get-Folder #aciona  a função que adquire o caminho que o novo arquivo será salvo
$xmldoc = [xml](gc $file -encoding UTF8) #Importa o arquivo XML para ser tratado pelo script

#Cria função que faz as tratativas dos valores dentro do XML
function tratativasXML {
 
#Tratamento de valores
$atacado = $produto.valor_atacado #alimenta o valor_atacado em uma váriavel
$dropshipping = $produto.valor_dropshipping #alimenta o valor_dropshipping em uma variável
$valor_atacado = $xmldoc.CreateElement("valor_atacado") #Cria elemento XML valor_atacado
$valor_atacado.set_InnerText($produto.valor_atacado) #insere o valor do elemento XML valor_atacado
$valor_dropshipping = $xmldoc.CreateElement("valor_dropshipping") #Cria elemento XML valor_dropshipping
$valor_dropshipping.set_InnerText($produto.valor_dropshipping) #insere o valor do elemento valor_dropshipping

#Tratamento de dimensões
$dimensoes = $produto.dimensao_caixa_cm #adquire as dimensões originais que precisam ser alteradas
$dimensoes = $dimensoes -replace '\s','' #deleta os espaços em branco
$CharArray = $dimensoes.Split("x") #corta os três valores utilizando x como referencia
$dim_altura = $CharArray[2] #cria o valor dim_altura
$dim_largura = $CharArray[0] #cria o valor dim_largura
$dim_comprimento = $CharArray[1] #cria o valor dim_comprimento
$dimensao_altura = $xmldoc.CreateElement("dimensao_altura") #Cria elemento xml dimensao_altura
$dimensao_altura.set_InnerText($dim_altura) #insere o valor do elemento XML dimensao_altura
$dimensao_largura = $xmldoc.CreateElement("dimensao_largura") #Cria o elemento XML dimensao_largura
$dimensao_largura.set_InnerText($dim_largura) #insere o valor do elemento dimensao_largura
$dimensao_comprimento = $xmldoc.CreateElement("dimensao_comprimento") #Cria o elemento XML dimensao_comprimento
$dimensao_comprimento.set_InnerText($dim_comprimento) #insere o valor do elemento dimensao_comprimento

#Tratativa de peso
$peso_gramas = $produto.peso_gramas #Adquire o peso do produto em gramas
$peso_gramas = $peso_gramas /1000 #Divide o peso por mil
$peso_gramas = $peso_gramas-replace ',','.' #substitui o separador padrão vírgula (,) por ponto (.)
$p_gramas = $xmldoc.CreateElement("peso_gramas") #Cria o elemento XML peso_gramas
$p_gramas.set_InnerText($peso_gramas) #Insere o valor já calculado de peso_gramas

#Esse bloco abaixo indexa todos os valores criados anteriormente dentro do XML
$estoque.AppendChild($valor_atacado) |Out-Null 
$estoque.AppendChild($valor_dropshipping)  |Out-Null
$estoque.AppendChild($dimensao_altura)  |Out-Null
$estoque.AppendChild($dimensao_largura)  |Out-Null
$estoque.AppendChild($dimensao_comprimento)  |Out-Null
$estoque.AppendChild($p_gramas)  |Out-Null


}


foreach ($produto in $xmlDoc.dados.produto) #Faz um condicional que passa por cada produto dentro da lista de produtos executando a função abaixo.
{
    foreach ($estoque in $produto.estoque) #Faz um condicional que passa por cada item de estoque dentro de cada produto.
    {
       tratativasXML #Chama a função que inclui tudo que foi programado dentro da função tratativasXML
    }
}
$data=Get-Date -format "yyyyMMdd" #adquire data para o nome do arquivo
$xmldoc.Save("$folder\XMLmodificado-$data.xml") #salva o arquivo modificado
write-output "Seu arquivo foi salvo em $folder\XMLmodificado.xml"
pause
