
# Limpa o terminal para uma exibição limpa.
Clear-Host

# Define a codificação do console para UTF-8 para suportar caracteres especiais.
[System.Console]::OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host "========================================="
Write-Host "   Gerenciador de Impressoras"
Write-Host "========================================="
Write-Host

# --- 1. Listar Impressoras ---
Write-Host "Procurando impressoras instaladas..."
try {
    # Obtém todas as impressoras disponíveis no sistema.
    $printers = Get-Printer | Select-Object -ExpandProperty Name
} catch {
    Write-Host -ForegroundColor Red "Erro ao buscar impressoras: $_"
    Write-Host "Pressione Enter para sair."
    Read-Host
    exit
}


if (-not $printers) {
    Write-Host -ForegroundColor Yellow "Nenhuma impressora encontrada."
    Write-Host "Pressione Enter para sair."
    Read-Host
    exit
}

Write-Host -ForegroundColor Green "Impressoras encontradas:"
$i = 1
foreach ($printer in $printers) {
    Write-Host "$i. $($printer.Name)"
    $i++
}
Write-Host "0. Cancelar e Sair"
Write-Host

# --- 2. Obter a Escolha do Usuário ---
$choice = -1
while ($choice -lt 0 -or $choice -gt $printers.Count) {
    try {
        $choice = Read-Host "Digite o número da impressora que deseja reinstalar (ou 0 para sair)"
        if ($choice -lt 0 -or $choice -gt $printers.Count) {
            Write-Host -ForegroundColor Yellow "Opção inválida. Por favor, tente novamente."
        }
    } catch {
        Write-Host -ForegroundColor Yellow "Entrada inválida. Por favor, digite um número."
    }
}

if ($choice -eq 0) {
    Write-Host "Operação cancelada."
    exit
}

# --- 3. Preparar para Reinstalação ---
# O índice do array é a escolha do usuário menos 1.
$selectedPrinter = $printers[$choice - 1]

# Salva os detalhes da impressora antes de removê-la.
$printerName = $selectedPrinter.Name
$printerDriver = $selectedPrinter.DriverName
$printerPort = $selectedPrinter.PortName

Write-Host
Write-Host -ForegroundColor Cyan "Você selecionou a impressora: $printerName"
Write-Host -ForegroundColor Cyan "   Driver: $printerDriver"
Write-Host -ForegroundColor Cyan "   Porta: $printerPort"
Write-Host

# Confirmação final
$confirmation = Read-Host "Tem certeza que deseja REMOVER e REINSTALAR esta impressora? (s/n)"

if ($confirmation -ne 's') {
    Write-Host "Operação cancelada pelo usuário."
    Write-Host "Pressione Enter para sair."
    Read-Host
    exit
}

# --- 4. Remover a Impressora ---
Write-Host
Write-Host -ForegroundColor Yellow "Removendo a impressora '$printerName'..."
try {
    Remove-Printer -Name $printerName -ErrorAction Stop
    Write-Host -ForegroundColor Green "Impressora removida com sucesso."
} catch {
    Write-Host -ForegroundColor Red "ERRO ao remover a impressora: $_"
    Write-Host "A reinstalação foi cancelada."
    Write-Host "Pressione Enter para sair."
    Read-Host
    exit
}

# --- 5. Reinstalar a Impressora ---
Write-Host
Write-Host -ForegroundColor Yellow "Reinstalando a impressora '$printerName'..."
try {
    Add-Printer -Name $printerName -DriverName $printerDriver -PortName $printerPort -ErrorAction Stop
    Write-Host -ForegroundColor Green "Impressora '$printerName' reinstalada com sucesso!"
} catch {
    Write-Host -ForegroundColor Red "ERRO ao reinstalar a impressora: $_"
    Write-Host -ForegroundColor Red "A impressora foi removida, mas não pôde ser reinstalada."
    Write-Host -ForegroundColor Red "Verifique o nome do driver e da porta e tente adicioná-la manualmente."
}

Write-Host
Write-Host "Operação concluída. Pressione Enter para sair."
Read-Host
