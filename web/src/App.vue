<script setup>
import { computed, ref } from 'vue'
import { ElMessage } from 'element-plus'
import { DocumentChecked, Download, InfoFilled, UploadFilled } from '@element-plus/icons-vue'

const sourceType = ref('outbound')
const selectedFiles = ref([])
const uploading = ref(false)
const resultLogPath = ref('')
const uploadRef = ref(null)

const sourceOptions = [
  {
    value: 'outbound',
    title: '出库表',
    description: '按硬件产品信息映射货品名称，生成出库发货登记表。',
    accept: '.xls,.xlsx',
  },
  {
    value: 'weidian',
    title: '微店表',
    description: '过滤已关闭订单，按商品 ID 汇总配件数量。',
    accept: '.xlsx',
  },
]

const currentSource = computed(() => sourceOptions.find((item) => item.value === sourceType.value))
const selectedFileName = computed(() => {
  if (selectedFiles.value.length === 0) return '尚未选择文件'
  if (selectedFiles.value.length === 1) return selectedFiles.value[0].name
  return `已选择 ${selectedFiles.value.length} 个文件，将打包下载 zip`
})

function handleFileChange(_uploadFile, uploadFiles) {
  selectedFiles.value = uploadFiles.map((item) => item.raw).filter(Boolean)
  resultLogPath.value = ''
}

function handleFileRemove(_uploadFile, uploadFiles) {
  selectedFiles.value = uploadFiles.map((item) => item.raw).filter(Boolean)
  resultLogPath.value = ''
}

function downloadFilename(response) {
  const disposition = response.headers.get('Content-Disposition') || ''
  const encodedMatch = disposition.match(/filename\*=UTF-8''([^;]+)/i)
  if (encodedMatch?.[1]) return decodeURIComponent(encodedMatch[1])
  const match = disposition.match(/filename="?([^";]+)"?/i)
  if (match?.[1]) return match[1]
  if (selectedFiles.value.length > 1) return '处理后文件.zip'
  return `处理后_${selectedFiles.value[0].name.replace(/\.[^.]+$/, '')}.xlsx`
}

async function transformFile() {
  if (selectedFiles.value.length === 0) {
    ElMessage.warning('请先选择 Excel 文件')
    return
  }

  const formData = new FormData()
  formData.append('sourceType', sourceType.value)
  selectedFiles.value.forEach((file) => {
    formData.append('files', file)
  })

  uploading.value = true
  try {
    const response = await fetch('/api/transform', {
      method: 'POST',
      body: formData,
    })
    if (!response.ok) {
      throw new Error(await response.text())
    }

    resultLogPath.value = response.headers.get('X-Transform-Log-Path') || ''
    const blob = await response.blob()
    const url = URL.createObjectURL(blob)
    const link = document.createElement('a')
    link.href = url
    link.download = downloadFilename(response)
    link.click()
    URL.revokeObjectURL(url)
    ElMessage.success(selectedFiles.value.length > 1 ? '批量转换完成，已下载 zip' : '转换完成')
    selectedFiles.value = []
    uploadRef.value?.clearFiles()
  } catch (error) {
    ElMessage.error(error.message || '转换失败')
  } finally {
    uploading.value = false
  }
}
</script>

<template>
  <main class="page">
    <section class="shell">
      <aside class="hero-panel">
        <div class="brand-badge">
          <el-icon>
            <DocumentChecked />
          </el-icon>
          <span>TableFusion</span>
        </div>
        <h1>表格转换工具</h1>
        <p class="hero-desc">
          处理出库表与微店表，转换完成立即下载标准表格
        </p>

        <div class="feature-list">
          <div class="feature-item">
            <span>01</span>
            <strong>选择类型</strong>
            <p>出库表 / 微店表使用独立转换规则。</p>
          </div>
          <div class="feature-item">
            <span>02</span>
            <strong>上传 Excel</strong>
            <p>文件只在本机内存中处理，不落库存储。</p>
          </div>
          <div class="feature-item">
            <span>03</span>
            <strong>下载结果</strong>
            <p>单文件直接下载，多文件自动打包 zip。</p>
          </div>
        </div>
      </aside>

      <el-card class="converter-card" shadow="never">
        <div class="card-header">
          <div>
            <p class="eyebrow">Excel Converter</p>
            <h2>开始转换</h2>
          </div>
          <el-tag type="success" effect="light">本地运行</el-tag>
        </div>

        <el-alert class="tip-alert" type="info" :closable="false" show-icon>
          <template #title>
            单文件会直接下载 xlsx；多个文件会统一转换并打包 zip 下载。
          </template>
        </el-alert>

        <el-form class="converter-form" label-position="top">
          <el-form-item label="表格类型">
            <div class="source-grid">
              <button v-for="item in sourceOptions" :key="item.value" class="source-card"
                :class="{ active: sourceType === item.value }" type="button"
                @click="sourceType = item.value; resultLogPath = ''">
                <strong>{{ item.title }}</strong>
                <span>{{ item.description }}</span>
              </button>
            </div>
          </el-form-item>

          <el-form-item label="Excel 文件">
            <el-upload ref="uploadRef" class="upload-box" drag multiple :accept="currentSource.accept"
              :auto-upload="false" :on-change="handleFileChange" :on-remove="handleFileRemove">
              <el-icon class="upload-icon">
                <UploadFilled />
              </el-icon>
              <div class="upload-text">
                <strong>拖拽文件到这里，或点击选择</strong>
                <span>{{ currentSource.title }}支持：{{ currentSource.accept }}，可多选批量转换</span>
              </div>
            </el-upload>
          </el-form-item>

          <div class="file-state">
            <div>
              <span>当前文件</span>
              <strong>{{ selectedFileName }}</strong>
            </div>
            <el-icon>
              <InfoFilled />
            </el-icon>
          </div>

          <el-button class="submit-button" type="primary" size="large" :icon="Download" :loading="uploading"
            @click="transformFile">
            转换并下载
          </el-button>

          <el-alert v-if="resultLogPath" class="result-alert" type="success" :closable="false" show-icon>
            <template #title>转换完成</template>
          </el-alert>
        </el-form>
      </el-card>
    </section>
  </main>
</template>
