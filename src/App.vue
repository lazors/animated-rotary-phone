<template>
  <div class="container">
    <header>
      <h1>📊 XLSX Parser</h1>
      <p>Upload and parse Excel files to view data in JSON format</p>
    </header>

    <div class="upload-section">
      <div
        class="upload-area"
        :class="{ dragover: isDragOver }"
        @click="triggerFileInput"
        @dragover="handleDragOver"
        @dragleave="handleDragLeave"
        @drop="handleDrop"
      >
        <div class="upload-content">
          <div class="upload-icon">📁</div>
          <h3>Drop your XLSX file here</h3>
          <p>or <span class="browse-text">browse files</span></p>
          <input
            ref="fileInput"
            type="file"
            accept=".xlsx,.xls"
            style="display: none;"
            @change="handleFileSelect"
          >
        </div>
      </div>
    </div>

    <div v-if="showControls" class="controls">
      <div class="sheet-selector">
        <label for="sheetSelect">Select Sheet:</label>
        <select
          id="sheetSelect"
          v-model="selectedSheet"
        >
          <option value="">Choose a sheet...</option>
          <option
            v-for="sheet in sheetNames"
            :key="sheet"
            :value="sheet"
          >
            {{ sheet }}
          </option>
        </select>
      </div>
      <div class="action-buttons">
        <button
          class="btn btn-primary"
          @click="parseSelectedSheet"
        >
          Parse Data
        </button>
        <button
          v-if="showDownload"
          class="btn btn-secondary"
          @click="downloadJSON"
        >
          Download JSON
        </button>
        <button
          class="btn btn-danger"
          @click="clearAll"
        >
          Clear
        </button>
      </div>
    </div>

    <div v-if="showResults" class="results-section">
      <div class="results-header">
        <h3>Parsed Data</h3>
        <div class="results-info">
          <span>{{ recordCount }} records</span>
          <span>{{ fileName }}</span>
        </div>
      </div>
      <div class="results-content">
        <div class="data-table">
          <table v-if="currentData && currentData.length > 0">
            <thead>
              <tr>
                <th v-for="header in headers" :key="header">
                  {{ header }}
                </th>
              </tr>
            </thead>
            <tbody>
              <tr v-for="(row, index) in currentData" :key="index">
                <td v-for="header in headers" :key="header">
                  {{ row[header] || '' }}
                </td>
              </tr>
            </tbody>
          </table>
          <p v-else>No data found in the selected sheet.</p>
        </div>
        <div class="json-output">
          <h4>JSON Output</h4>
          <pre>{{ jsonOutput }}</pre>
        </div>
      </div>
    </div>

    <div v-if="errorMessage" class="error-message">
      <div class="error-content">
        <span class="error-icon">⚠️</span>
        <span class="error-text">{{ errorMessage }}</span>
        <button class="error-close" @click="hideError">&times;</button>
      </div>
    </div>
  </div>
</template>

<script>
import { ref, computed, onMounted } from 'vue'
import * as XLSX from 'xlsx'

export default {
  name: 'App',
  setup() {
    const workbook = ref(null)
    const currentData = ref(null)
    const fileName = ref('')
    const selectedSheet = ref('')
    const sheetNames = ref([])
    const isDragOver = ref(false)
    const errorMessage = ref('')
    const fileInput = ref(null)

    const showControls = computed(() => workbook.value !== null)
    const showResults = computed(() => currentData.value !== null)
    const showDownload = computed(() => currentData.value !== null)
    const recordCount = computed(() =>
      currentData.value ? `${currentData.value.length}` : '0'
    )

    const headers = computed(() => {
      if (!currentData.value || currentData.value.length === 0) return []

      const allKeys = new Set()
      currentData.value.forEach(row => {
        Object.keys(row).forEach(key => allKeys.add(key))
      })
      return Array.from(allKeys)
    })

    const jsonOutput = computed(() =>
      currentData.value ? JSON.stringify(currentData.value, null, 2) : ''
    )

    const triggerFileInput = () => {
      fileInput.value.click()
    }

    const handleDragOver = (e) => {
      e.preventDefault()
      isDragOver.value = true
    }

    const handleDragLeave = (e) => {
      e.preventDefault()
      isDragOver.value = false
    }

    const handleDrop = (e) => {
      e.preventDefault()
      isDragOver.value = false
      const files = e.dataTransfer.files
      if (files.length > 0) {
        processFile(files[0])
      }
    }

    const handleFileSelect = (e) => {
      const file = e.target.files[0]
      if (file) {
        processFile(file)
      }
    }

    const processFile = (file) => {
      if (!file) return

      const validTypes = [
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'application/vnd.ms-excel'
      ]

      if (!validTypes.includes(file.type) && !file.name.match(/\.(xlsx|xls)$/i)) {
        showError('Please select a valid Excel file (.xlsx or .xls)')
        return
      }

      fileName.value = file.name

      const reader = new FileReader()
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result)
          workbook.value = XLSX.read(data, { type: 'array' })
          populateSheetSelector()
          hideError()
        } catch (error) {
          showError('Error reading Excel file: ' + error.message)
        }
      }
      reader.readAsArrayBuffer(file)
    }

    const populateSheetSelector = () => {
      sheetNames.value = workbook.value.SheetNames

      if (workbook.value.SheetNames.length === 1) {
        selectedSheet.value = workbook.value.SheetNames[0]
      } else {
        selectedSheet.value = ''
      }
    }

    const parseSelectedSheet = () => {
      if (!selectedSheet.value) {
        showError('Please select a sheet to parse')
        return
      }

      try {
        const worksheet = workbook.value.Sheets[selectedSheet.value]
        currentData.value = XLSX.utils.sheet_to_json(worksheet)
        hideError()
      } catch (error) {
        showError('Error parsing sheet: ' + error.message)
      }
    }

    const downloadJSON = () => {
      if (!currentData.value) return

      const jsonStr = JSON.stringify(currentData.value, null, 2)
      const blob = new Blob([jsonStr], { type: 'application/json' })
      const url = URL.createObjectURL(blob)

      const a = document.createElement('a')
      a.href = url
      a.download = fileName.value.replace(/\.(xlsx|xls)$/i, '.json')
      document.body.appendChild(a)
      a.click()
      document.body.removeChild(a)
      URL.revokeObjectURL(url)
    }

    const clearAll = () => {
      workbook.value = null
      currentData.value = null
      fileName.value = ''
      selectedSheet.value = ''
      sheetNames.value = []
      fileInput.value.value = ''
      hideError()
    }

    const showError = (message) => {
      errorMessage.value = message
      setTimeout(() => {
        hideError()
      }, 5000)
    }

    const hideError = () => {
      errorMessage.value = ''
    }

    return {
      workbook,
      currentData,
      fileName,
      selectedSheet,
      sheetNames,
      isDragOver,
      errorMessage,
      fileInput,
      showControls,
      showResults,
      showDownload,
      recordCount,
      headers,
      jsonOutput,
      triggerFileInput,
      handleDragOver,
      handleDragLeave,
      handleDrop,
      handleFileSelect,
      parseSelectedSheet,
      downloadJSON,
      clearAll,
      hideError
    }
  }
}
</script>