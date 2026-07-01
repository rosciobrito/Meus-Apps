# Novo Código Apps Script — doPost com Salvamento de Imagem

> **Este código deve substituir o arquivo atual no Google Apps Script.**
> Ele inclui uma função `autorizarPermissoes()` específica para resolver o erro de permissão do Google Drive!

## 🚨 COMO CORRIGIR O ERRO DE PERMISSÃO DO DRIVE (Passo a Passo)

Quando adicionamos o comando `DriveApp` para salvar imagens, o Google exige que você autorize o acesso ao seu Google Drive **manualmente no editor** antes que o aplicativo consiga enviar fotos.

Faça exatamente o seguinte no editor do Apps Script:

1. Abra a planilha e vá em **Extensões → Apps Script**.
2. Cole todo o código abaixo no seu arquivo `Código.gs` e clique em **Salvar** (💾).
3. No painel superior, onde tem um menu suspenso ao lado do botão **▶ Executar**, selecione a função **`autorizarPermissoes`**.
4. Clique no botão **▶ Executar**.
5. Vai aparecer uma janela dizendo **"Autorização necessária"**:
   * Clique em **Revisar permissões**.
   * Escolha a sua conta do Google.
   * Se aparecer um aviso *"O Google não verificou este app"*, clique em **Avançado** (texto pequeno) e depois em **Acessar Projeto sem título (não seguro)**.
   * Clique em **Permitir**.
6. Agora vá em **Implantar → Gerenciar implantações**.
7. Clique no ícone de **lápis (editar)** na implantação ativa.
8. Em "Versão", selecione **"Nova versão"** e clique em **Implantar**.

Pronto! A permissão para o Google Drive estará concedida e as imagens serão salvas perfeitamente sem erros!

---

## Código Completo para Colar no Apps Script

```javascript
/**
 * ⚠️ RODE ESTA FUNÇÃO UMA VEZ NO EDITOR PARA AUTORIZAR O GOOGLE DRIVE!
 * Selecione "autorizarPermissoes" no topo e clique em ▶ Executar.
 */
function autorizarPermissoes() {
  DriveApp.getRootFolder();
  SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("✅ Permissões de Planilha e Google Drive autorizadas com sucesso!");
}

function doPost(e) {
  try {
    // Pega a aba ativa da planilha atual
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // Converte os dados recebidos do aplicativo (que vêm em JSON)
    var data = JSON.parse(e.postData.contents);
    
    // Validação de Token de segurança
    var token_esperado = "inventario2026"; 
    
    if (token_esperado !== "" && data._token !== token_esperado) {
       throw new Error("Acesso negado: Token inválido");
    }

    // Remove o token antes de salvar
    delete data._token;

    // ═══════════════════════════════════════════
    // SALVAR IMAGEM NO GOOGLE DRIVE
    // ═══════════════════════════════════════════
    var etiquetaPath = "";
    
    if (data._imagem) {
      var imageBase64 = data._imagem;
      delete data._imagem;  // Remove do payload antes de salvar na planilha
      
      try {
        // Encontrar a pasta pai da planilha no Drive
        var spreadsheetFile = DriveApp.getFileById(
          SpreadsheetApp.getActiveSpreadsheet().getId()
        );
        var parentFolder = spreadsheetFile.getParents().next();
        
        // Criar ou encontrar a pasta "Aparelhos_Images"
        var folder;
        var folders = parentFolder.getFoldersByName("Aparelhos_Images");
        if (folders.hasNext()) {
          folder = folders.next();
        } else {
          folder = parentFolder.createFolder("Aparelhos_Images");
        }
        
        // Gerar nome do arquivo: IMEI.Etiqueta.HHMMSS.jpg
        var now = new Date();
        var timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), "HHmmss");
        var imeiPart = data.imei || "sem_imei";
        // Limpar caracteres inválidos do IMEI (pode ter barras se tiver 2 IMEIs)
        imeiPart = imeiPart.replace(/[\/\\|]/g, "_");
        var fileName = imeiPart + ".Etiqueta." + timestamp + ".jpg";
        
        // Decodificar base64 e criar o arquivo
        var blob = Utilities.newBlob(
          Utilities.base64Decode(imageBase64),
          "image/jpeg",
          fileName
        );
        var file = folder.createFile(blob);
        
        // Caminho relativo (mesmo padrão dos dados existentes)
        etiquetaPath = "Aparelhos_Images/" + fileName;
        
      } catch (imgError) {
        // Se falhar ao salvar imagem, continua salvando os dados
        etiquetaPath = "ERRO: " + imgError.toString();
      }
    }

    // ═══════════════════════════════════════════
    // SALVAR DADOS NA PLANILHA
    // ═══════════════════════════════════════════
    // Coluna A (Código) = vazio — é preenchida por fórmula da planilha
    // Colunas B-I = dados do app
    // Coluna J (flag) = FALSE (padrão)
    // Coluna K (Etiqueta) = caminho da imagem salva no Drive
    sheet.appendRow([
      "",                        // A: Código (fórmula da planilha)
      new Date(),                // B: Data/Hora do registro
      data.tipo || "",           // C: Tipo
      data.marca || "",          // D: Marca
      data.modelo || "",         // E: Modelo
      data.imei || "",           // F: IMEI
      data.sn || "",             // G: Nº Série
      data.acessorios || "",     // H: Acessórios
      data.material || "",       // I: Material
      false,                     // J: Flag (padrão FALSE)
      etiquetaPath               // K: Etiqueta (caminho da imagem no Drive)
    ]);

    // Retorna mensagem de sucesso
    return ContentService.createTextOutput(JSON.stringify({ 
      "status": "success",
      "etiqueta": etiquetaPath
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    // Em caso de erro, retorna a mensagem do erro
    return ContentService.createTextOutput(JSON.stringify({ 
      "status": "error", 
      "message": error.toString() 
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
```

---

## Como Funciona o Salvamento de Imagem

```mermaid
sequenceDiagram
    participant App as 📱 App Web
    participant GAS as ⚡ Apps Script
    participant Drive as 📁 Google Drive
    participant Sheet as 📊 Planilha

    App->>GAS: POST {dados + _imagem (base64)}
    GAS->>GAS: Validar token
    GAS->>GAS: Decodificar base64 → Blob JPEG
    GAS->>Drive: Salvar em Aparelhos_Images/IMEI.Etiqueta.HHMMSS.jpg
    Drive-->>GAS: Arquivo criado ✅
    GAS->>Sheet: appendRow (dados + caminho da imagem)
    GAS-->>App: { status: "success" }
```

---

## Estrutura da Planilha (11 colunas)

| Coluna | Campo | Fonte |
|:--|:--|:--|
| A | Código | Fórmula da planilha |
| B | Data/Hora | `new Date()` |
| C | Tipo | IA + edição |
| D | Marca | IA + edição |
| E | Modelo | IA + edição |
| F | IMEI | IA + edição |
| G | Nº Série | IA + edição |
| H | Acessórios | IA + edição |
| I | Material | IA + edição |
| J | Flag | `false` (padrão) |
| **K** | **Etiqueta** | **Caminho da imagem salva no Drive** |

---

## Padrão de Nomes das Imagens

O arquivo é salvo seguindo o mesmo padrão dos dados existentes:

```
Aparelhos_Images/358459768232759.Etiqueta.143510.jpg
                 └─── IMEI ────┘          └ HH:MM:SS
```

- Se houver 2 IMEIs (separados por `/`), as barras são substituídas por `_`
- Se não houver IMEI, usa `sem_imei`

---

## ⚠️ Notas Importantes

1. **Imagem reduzida**: O app envia uma versão redimensionada (~1200px de largura, JPEG 80%) para economia de dados. A imagem original em alta resolução é usada apenas para a IA.

2. **Pasta no Drive**: As imagens são salvas na pasta `Aparelhos_Images` que fica no **mesmo diretório** da planilha no Google Drive. Se a pasta não existir, ela é criada automaticamente.

3. **Permissões**: Na primeira execução, o Apps Script pedirá permissão para acessar o Drive. Aceite.

4. **Tolerância a falhas**: Se a imagem falhar ao salvar (ex: muito grande), os dados textuais ainda são salvos na planilha normalmente. A coluna Etiqueta mostrará `ERRO: [mensagem]`.


