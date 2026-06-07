<script setup lang="ts">
import * as docx from 'docx'
import JSZip from 'jszip'
import { ref } from 'vue'
import * as udf from './scripts/udf'

enum ToolEnum {
  toDocx,
  versionCheck,
}

const docxFile = ref<File | null>(null)
const docxConsole = ref<string>('')

const versionFile = ref<File | null>(null)
const versionConsole = ref<string>('')

async function processFile(event: Event, type: ToolEnum.toDocx | ToolEnum.versionCheck) {
  const input = event.target as HTMLInputElement
  if (!input.files || input.files.length === 0) {
    switch (type) {
      case ToolEnum.toDocx:
        docxFile.value = null
        docxConsole.value = ''
        break

      case ToolEnum.versionCheck:
        versionFile.value = null
        versionConsole.value = ''
        break
    }

    return
  }

  const file = input.files[0]

  switch (type) {
    case ToolEnum.toDocx:
      docxFile.value = file
      docxConsole.value = 'Dosya okunuyor...\n'
      break

    case ToolEnum.versionCheck:
      versionFile.value = file
      versionConsole.value = 'Dosya okunuyor...\n'
      break

    default:
      throw new Error(`Bilinmeyen (${type}) islem tipi.`)
  }

  try {
    const appendLog = (msg: string) => {
      switch (type) {
        case ToolEnum.toDocx:
          docxConsole.value += msg
          break
        case ToolEnum.versionCheck:
          versionConsole.value += msg
          break
      }
    }

    const xmlData = await extractXmlFromZipSync(file)
    if (!xmlData) {
      appendLog('Hata: UDF icinde content.xml bulunamadi veya dosya bos.\n')
      return
    }

    const udfParser = new udf.Document(xmlData)
    appendLog(`Dokuman versiyonu: ${udfParser.document.template.$?.format_id ?? 'bilinmiyor'}\n`)

    if (type === ToolEnum.toDocx) {
      const docxDocument = udfParser.toDocx()
      if (docxDocument) {
        appendLog('Dokuman basariyla donusturuldu.\n')
      }
      else {
        appendLog('Dokuman donusturucu icerik dondurmedi.\n')
        return
      }

      docxConsole.value += 'Donusturulmus dokuman aktariliyor...\n'
      const docxBlob = await docx.Packer.toBlob(docxDocument)
      downloadBlob(docxBlob, `${file.name}.docx`)
      docxConsole.value += 'Dokuman yuklenmeye hazir.\n'
    }
  }
  catch (error) {
    console.error('Dosya islenirken hata olustu:', error)
    const errorMsg = `Hata: ${error instanceof Error ? error.message : 'Islem sirasinda bilinmeyen hata'}.\n`
    switch (type) {
      case ToolEnum.toDocx:
        docxConsole.value += errorMsg
        break
      case ToolEnum.versionCheck:
        versionConsole.value += errorMsg
        break
    }
  }
}

async function extractXmlFromZipSync(file: File): Promise<string | undefined> {
  const zip = new JSZip()
  const loadedZip = await zip.loadAsync(file)
  return await loadedZip.file('content.xml')?.async('string')
}

function downloadBlob(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob)
  const link = document.createElement('a')
  link.href = url
  link.download = filename
  document.body.appendChild(link)
  link.click()
  document.body.removeChild(link)
  URL.revokeObjectURL(url)
}
</script>

<template>
  <div class="card-grid">
    <div class="card">
      <h2 class="card-title">
        Docx Dönüştürücü
      </h2>
      <p class="card-desc">
        UDF formatını DOCX'e dönüştürür.
      </p>

      <div class="input-group">
        <label for="xmlFileInput" class="file-label">Dosya:</label> <br>
        <input id="xmlFileInput" type="file" accept=".udf" @change="processFile($event, ToolEnum.toDocx)">
      </div>

      <div v-if="docxFile" class="info-section">
        <p class="file-label">
          Seçilen Dosya: <br> <strong>{{ docxFile.name }}</strong>
        </p>
      </div>

      <div v-if="docxConsole" class="info-section">
        <h3>Konsol Çıktısı:</h3>
        <pre class="console-box">{{ docxConsole }}</pre>
      </div>
    </div>

    <div class="card">
      <h2 class="card-title">
        Versiyon Kontrolü
      </h2>
      <p class="card-desc">
        UDF formatının versiyonunu kontrol etmeye yarar.
      </p>

      <div class="input-group">
        <label for="versionFileInput" class="file-label">Dosya:</label> <br>
        <input id="versionFileInput" type="file" accept=".udf" @change="processFile($event, ToolEnum.versionCheck)">
      </div>

      <div v-if="versionFile" class="info-section">
        <p class="file-label">
          Seçilen Dosya: <br> <strong>{{ versionFile.name }}</strong>
        </p>
      </div>

      <div v-if="versionConsole" class="info-section">
        <h3>Konsol Çıktısı:</h3>
        <pre class="console-box">{{ versionConsole }}</pre>
      </div>
    </div>
  </div>
</template>

<style scoped>
.page {
  width: 100%;
  max-width: 800px;
  margin: 0 auto;
}

.container {
  margin-bottom: 2rem;
  font-family: sans-serif;
}

.title,
.file-label {
  text-decoration: underline;
}

.input-group {
  margin: 1rem 0;
}

.info-section {
  margin-top: 1rem;
}

.console-box {
  background-color: #f4f4f4;
  border: 1px solid #ddd;
  color: #333;
  padding: 10px;
  max-height: 200px;
  overflow-y: auto;
  white-space: pre-wrap;
  word-break: break-all;
  font-family: monospace;
}
</style>
