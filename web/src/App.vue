<script setup>
import { ref } from 'vue'
import { ElMessage } from 'element-plus'

const sourceType = ref('outbound')
const selectedFile = ref(null)
const uploading = ref(false)

function beforeUpload(file) {
  selectedFile.value = file
  return false
}

async function transformFile() {
  if (!selectedFile.value) {
    ElMessage.warning('请先选择 Excel 文件')
    return
  }

  const formData = new FormData()
  formData.append('sourceType', sourceType.value)
  formData.append('file', selectedFile.value)

  uploading.value = true
  try {
    const response = await fetch('/api/transform', {
      method: 'POST',
      body: formData,
    })
    if (!response.ok) {
      throw new Error(await response.text())
    }

    const blob = await response.blob()
    const url = URL.createObjectURL(blob)
    const link = document.createElement('a')
    link.href = url
    link.download = `standard-${selectedFile.value.name}`
    link.click()
    URL.revokeObjectURL(url)
    ElMessage.success('转换完成')
  } catch (error) {
    ElMessage.error(error.message || '转换失败')
  } finally {
    uploading.value = false
  }
}
</script>

<template>
  <main class="page">
    <el-card class="converter-card" shadow="never">
      <template #header>
        <div class="card-header">
          <h1>表格转换工具</h1>
          <p>选择表格类型，上传 Excel，下载转换后的标准表格。</p>
        </div>
      </template>

      <el-form label-position="top">
        <el-form-item label="表格类型">
          <el-radio-group v-model="sourceType">
            <el-radio-button label="outbound">出库表</el-radio-button>
            <el-radio-button label="weidian">微店表</el-radio-button>
          </el-radio-group>
        </el-form-item>

        <el-form-item label="Excel 文件">
          <el-upload
            drag
            accept=".xlsx,.xls"
            :auto-upload="false"
            :limit="1"
            :before-upload="beforeUpload"
          >
            <div class="upload-text">
              <strong>拖拽文件到这里</strong>
              <span>或点击选择 Excel 文件</span>
            </div>
          </el-upload>
        </el-form-item>

        <el-button type="primary" size="large" :loading="uploading" @click="transformFile">
          转换并下载
        </el-button>
      </el-form>
    </el-card>
  </main>
</template>
