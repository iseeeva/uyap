<script setup lang="ts">
import * as docx from 'docx'
import { ref } from 'vue'
import * as udf from './scripts/udf'

const selectedFile = ref<File | null>(null)
const progressConsole = ref<string>('')

function handleFileChange(event: Event) {
  const input = event.target as HTMLInputElement
  if (input.files && input.files.length > 0) {
    const reader = new FileReader()
    reader.onload = async (e) => {
      const readerResult = e.target?.result as string

      try {
        const udfParser = new udf.Document(readerResult)
        const blob = udfParser.toDocx()

        if (blob) {
          progressConsole.value = 'Document parsed successfully.\n'
          const docxBlob = await docx.Packer.toBlob(blob)
          const url = URL.createObjectURL(docxBlob)
          const link = document.createElement('a')
          link.href = url
          link.download = 'parsed_document.docx'
          document.body.appendChild(link)
          link.click()
          document.body.removeChild(link)
          URL.revokeObjectURL(url)
        }
        else {
          progressConsole.value = 'Document parsing returned no content.\n'
        }
      }
      catch (error) {
        console.error('Error during parsing:', error)
        progressConsole.value = `Parsing error.\n`
      }
    }

    reader.onerror = (e) => {
      console.error('Error reading file:', e)
      progressConsole.value = 'File read error.\n'
    }

    selectedFile.value = input.files[0]
    reader.readAsText(selectedFile.value)
    console.log('Selected file name:', selectedFile.value.name)
  }
  else {
    selectedFile.value = null
    progressConsole.value = ''
    console.log('No file selected.')
  }
}
</script>

<template>
  <div>
    <label for="xmlFileInput">Select XML File:</label>
    <input id="xmlFileInput" type="file" accept=".xml" @change="handleFileChange">

    <div v-if="selectedFile">
      <p>Selected File: <strong>{{ selectedFile.name }}</strong></p>
    </div>
    <div v-if="progressConsole">
      <h3>File Content:</h3>
      <pre>{{ progressConsole }}</pre>
    </div>
  </div>
</template>

<style scoped>
pre {
  background-color: #f4f4f4;
  border: 1px solid #ddd;
  color: black;
  padding: 10px;
  max-height: 200px;
  overflow-y: auto;
  white-space: pre-wrap;
  word-break: break-all;
}
</style>
