<template>
  <div class="app-container">
    <header class="app-header">
      <h1>文件二维码工具 Web 版</h1>
      <div class="header-controls">
        <button @click="toggleTheme">切换主题</button>
        <button @click="resetConfig">重置配置</button>
      </div>
    </header>

    <main class="app-main">
      <div class="mode-selector">
        <button
          :class="{ active: mode === 'region' }"
          @click="mode = 'region'"
        >
          Excel 区域模式
        </button>
        <button
          :class="{ active: mode === 'file' }"
          @click="mode = 'file'"
        >
          文件模式
        </button>
      </div>

      <div class="control-panel">
        <!-- 文件上传 -->
        <div class="file-upload">
          <h2>{{ mode === 'region' ? '上传Excel文件' : '上传任意文件' }}</h2>
          <input type="file" @change="handleFileUpload" ref="fileInput" accept=".xlsx,.xlsm,.xltx,.xltm">
          <div v-if="uploadedFile" class="file-preview">
            <span>{{ uploadedFile.name }}</span>
            <button @click="removeFile">移除</button>
          </div>
        </div>

        <!-- Excel 区域设置 -->
        <div v-if="mode === 'region'" class="excel-settings">
          <div class="form-group">
            <label>区域坐标:</label>
            <input type="text" v-model="region" placeholder="例如: A1:D10">
          </div>
          <div class="form-group">
            <label>选择 Sheet:</label>
            <select v-model="sheetName">
              <option v-for="sheet in sheets" :key="sheet" :value="sheet">
                {{ sheet }}
              </option>
            </select>
          </div>
        </div>

        <!-- 二维码设置 -->
        <div class="qr-settings">
          <div class="form-group">
            <label>数据块大小 (字节):</label>
            <select v-model="maxChunkSize">
              <option value="1200">1200</option>
              <option value="1500">1500</option>
              <option value="1800">1800</option>
              <option value="2000">2000</option>
              <option value="2500">2500</option>
            </select>
          </div>
          <div class="form-group">
            <label>数据版本:</label>
            <span>{{ version }}</span>
          </div>
        </div>

        <!-- 操作按钮 -->
        <div class="action-buttons">
          <button @click="serialize" :disabled="!uploadedFile">序列化</button>
          <button @click="generateQR" :disabled="!sessionId">生成二维码</button>
          <button @click="scanImages">图片恢复</button>
          <button @click="scanVideo">视频恢复</button>
        </div>

        <!-- 进度显示 -->
        <div class="progress-section">
          <div class="progress-bar">
            <div class="progress-fill" :style="{ width: progress + '%' }"></div>
          </div>
          <div class="progress-label">{{ progressMessage }}</div>
        </div>
      </div>

      <!-- 二维码预览 -->
      <div v-if="qrPreviews.length" class="qr-preview-section">
        <h2>二维码预览 ({{ qrPreviews.length }} 个中的前 {{ qrPreviews.length > 3 ? 3 : qrPreviews.length }} 个)</h2>
        <div class="qr-grid">
          <div v-for="(qr, index) in qrPreviews" :key="index" class="qr-card">
<!--            <img :src="qr.data" alt="QR Code">-->
            <img :src="`data:image/png;base64,${qr.data}`" alt="QR Code">

            <div class="qr-name">{{ qr.name }}</div>
          </div>
        </div>
        <button @click="downloadQR">下载所有二维码</button>
      </div>

      <!-- 日志面板 -->
      <div class="log-section">
        <div class="log-header">
          <h2>操作记录</h2>
          <button @click="clearLogs">清空日志</button>
        </div>
        <div class="log-content">
          <div v-for="(entry, index) in logEntries" :key="index" class="log-entry">
            {{ entry }}
          </div>
        </div>
      </div>
    </main>

    <!-- 文件恢复结果 -->
    <div v-if="restoredFile" class="restore-result">
      <h2>文件恢复成功!</h2>
      <p>{{ restoredFileName }}</p>
      <button @click="downloadRestoredFile">下载恢复的文件</button>
    </div>
  </div>
</template>

<script>
import axios from 'axios';

export default {
  data() {
    return {
      mode: 'region', // 'region' 或 'file'
      uploadedFile: null,
      region: 'A1:D10',
      sheetName: '',
      sheets: [],
      maxChunkSize: 1800,
      version: 8,
      sessionId: '',
      progress: 0,
      progressMessage: '就绪',
      qrPreviews: [],
      logEntries: [
        '=== 文件二维码工具 Web 版 ===',
        '就绪，请选择操作'
      ],
      restoredFile: null,
      restoredFileName: ''
    };
  },
  methods: {
    async handleFileUpload(event) {
  const file = event.target.files[0];
  if (!file) return;

  // 验证文件扩展名
  const validExtensions = ['.xlsx', '.xlsm', '.xltx', '.xltm'];
  if (!validExtensions.some(ext => file.name.toLowerCase().endsWith(ext))) {
    this.addLog('错误: 仅支持 .xlsx, .xlsm, .xltx, .xltm 格式的Excel文件');
    return;
  }

  this.uploadedFile = file;
  this.addLog(`已上传文件: ${file.name}`);

  try {
    const formData = new FormData();
    formData.append('file', file);

    const response = await axios.post('/api/get-sheets', formData, {
      headers: {
        'Content-Type': 'multipart/form-data'
      }
    });

    this.sheets = response.data.sheets;
    this.sheetName = this.sheets[0] || '';
    this.addLog(`加载成功，找到 ${this.sheets.length} 个工作表`);

  } catch (error) {
    const errorMsg = error.response?.data?.detail ||
                    error.response?.data?.message ||
                    error.message;
    this.addLog(`加载失败: ${errorMsg}`);
  }
}

,
    removeFile() {
      this.uploadedFile = null;
      this.$refs.fileInput.value = '';
      this.addLog('已移除上传的文件');
    },
    async serialize() {
      if (!this.uploadedFile) {
        this.addLog('错误: 请先上传文件');
        return;
      }

      try {
        this.progress = 0;
        this.progressMessage = '开始序列化...';

        // 读取文件内容为Base64
        const reader = new FileReader();
        reader.onload = async () => {
          const base64Data = reader.result;

          const request = {
            mode: this.mode,
            file_path: base64Data,
            version: this.version,
            max_chunk_size: this.maxChunkSize
          };

          if (this.mode === 'region') {
            request.region = this.region;
            request.sheet_name = this.sheetName;
          }

          try {
            const response = await axios.post('/api/serialize', request);
            this.sessionId = response.data.session_id;
            this.addLog(`序列化成功，数据大小: ${response.data.data_size} 字节`);

            // 开始轮询进度
            this.pollSessionStatus();
          } catch (error) {
            this.addLog(`序列化失败: ${error.response?.data?.detail || error.message}`);
          }
        };

        reader.readAsDataURL(this.uploadedFile);
      } catch (error) {
        this.addLog(`序列化错误: ${error.message}`);
      }
    },
    async generateQR() {
      if (!this.sessionId) {
        this.addLog('错误: 请先完成序列化');
        return;
      }

      try {
        this.progress = 0;
        this.progressMessage = '开始生成二维码...';

        const request = {
          session_id: this.sessionId,
          max_chunk_size: this.maxChunkSize
        };

        const response = await axios.post('/api/generate-qr', request);
        this.qrPreviews = response.data.previews;
        this.addLog(`成功生成 ${response.data.count} 个二维码`);

        // 继续轮询进度直到完成
        this.pollSessionStatus();
      } catch (error) {
        this.addLog(`生成二维码失败: ${error.response?.data?.detail || error.message}`);
      }
    },
    async scanImages() {
      const input = document.createElement('input');
      input.type = 'file';
      input.multiple = true;
      input.accept = 'image/*';

      input.onchange = async (e) => {
        const files = Array.from(e.target.files);
        if (files.length === 0) return;

        this.progress = 0;
        this.progressMessage = '开始扫描图片...';
        this.addLog(`选择了 ${files.length} 张图片进行扫描`);

        try {
          // 修复：简化 Promise 链，避免语法错误
          const base64Files = await Promise.all(files.map(file => {
            return new Promise((resolve) => {
              const reader = new FileReader();
              reader.onload = () => resolve(reader.result);
              reader.readAsDataURL(file);
            });
          }));

          const request = {
            files: base64Files
          };

          const response = await axios.post('/api/scan-images', request);
          this.sessionId = response.data.session_id;
          this.addLog('开始扫描恢复...');

          // 开始轮询进度
          this.pollSessionStatus();
        } catch (error) {
          this.addLog(`图片恢复失败: ${error.response?.data?.detail || error.message}`);
        }
      };

      input.click();
    },
    async scanVideo() {
      const input = document.createElement('input');
      input.type = 'file';
      input.accept = 'video/*';

      input.onchange = async (e) => {
        const file = e.target.files[0];
        if (!file) return;

        this.progress = 0;
        this.progressMessage = '开始扫描视频...';
        this.addLog(`选择了视频文件: ${file.name}`);

        try {
          // 读取视频为Base64
          const base64Video = await new Promise((resolve) => {
            const reader = new FileReader();
            reader.onload = () => resolve(reader.result);
            reader.readAsDataURL(file);
          });

          const request = {
            video: base64Video
          };

          const response = await axios.post('/api/scan-video', request);
          this.sessionId = response.data.session_id;
          this.addLog('开始视频扫描...');

          // 开始轮询进度
          this.pollSessionStatus();
        } catch (error) {
          this.addLog(`视频恢复失败: ${error.response?.data?.detail || error.message}`);
        }
      };

      input.click();
    },
    async pollSessionStatus() {
      if (!this.sessionId) return;

      try {
        const response = await axios.get(`/api/session/${this.sessionId}`);
        const session = response.data;

        this.progress = session.progress;
        this.progressMessage = session.message;

        if (session.status === 'processing') {
          // 继续轮询
          setTimeout(() => this.pollSessionStatus(), 1000);
        } else if (session.status === 'completed') {
          this.addLog(session.message);

          // 如果有恢复的文件
          if (session.restored_file) {
            this.restoredFile = session.restored_file;
            this.restoredFileName = session.file_name;
            this.addLog(`文件恢复成功: ${session.file_name}`);
          }
        } else if (session.status === 'error') {
          this.addLog(`错误: ${session.message}`);
        }
      } catch (error) {
        this.addLog(`轮询会话状态失败: ${error.message}`);
      }
    },
    async downloadQR() {
      if (!this.sessionId) return;

      try {
        window.open(`/api/download/${this.sessionId}?file_type=qr`, '_blank');
        this.addLog('开始下载二维码文件');
      } catch (error) {
        this.addLog(`下载失败: ${error.message}`);
      }
    },
    async downloadRestoredFile() {
      if (!this.sessionId || !this.restoredFile) return;

      try {
        window.open(`/api/download/${this.sessionId}?file_type=restored`, '_blank');
        this.addLog(`下载恢复的文件: ${this.restoredFileName}`);
      } catch (error) {
        this.addLog(`下载失败: ${error.message}`);
      }
    },
    addLog(message) {
      const timestamp = new Date().toLocaleTimeString();
      this.logEntries.push(`[${timestamp}] ${message}`);
      // 保持日志在100条以内
      if (this.logEntries.length > 100) {
        this.logEntries.shift();
      }
    },
    clearLogs() {
      this.logEntries = [
        `[${new Date().toLocaleTimeString()}] 日志已清空`
      ];
    },
    toggleTheme() {
      document.body.classList.toggle('dark-theme');
      this.addLog('切换主题');
    },
    async resetConfig() {
      try {
        await axios.post('/api/reset-config');
        this.addLog('配置已重置为默认值');
      } catch (error) {
        this.addLog(`重置配置失败: ${error.message}`);
      }
    }
  }
};
</script>

<style>
:root {
  --primary-color: #3498db;
  --secondary-color: #2ecc71;
  --background-color: #f5f7fa;
  --card-bg: #ffffff;
  --text-color: #333333;
  --border-color: #e0e0e0;
  --progress-bg: #e0e0e0;
  --progress-fill: var(--primary-color);
}

.dark-theme {
  --background-color: #1e1e1e;
  --card-bg: #2d2d2d;
  --text-color: #f0f0f0;
  --border-color: #444444;
  --progress-bg: #444444;
}

* {
  box-sizing: border-box;
  margin: 0;
  padding: 0;
}

body {
  font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
  background-color: var(--background-color);
  color: var(--text-color);
  line-height: 1.6;
}

.app-container {
  max-width: 1400px;
  margin: 0 auto;
  padding: 20px;
}

.app-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 15px 0;
  border-bottom: 1px solid var(--border-color);
  margin-bottom: 20px;
}

.app-header h1 {
  font-size: 24px;
  color: var(--primary-color);
}

.header-controls button {
  background: var(--primary-color);
  color: white;
  border: none;
  padding: 8px 15px;
  border-radius: 4px;
  cursor: pointer;
  margin-left: 10px;
  transition: background 0.3s;
}

.header-controls button:hover {
  background: #2980b9;
}

.mode-selector {
  display: flex;
  margin-bottom: 20px;
  border-radius: 6px;
  overflow: hidden;
  box-shadow: 0 2px 5px rgba(0,0,0,0.1);
}

.mode-selector button {
  flex: 1;
  padding: 12px;
  border: none;
  background: var(--card-bg);
  color: var(--text-color);
  cursor: pointer;
  font-size: 16px;
  transition: all 0.3s;
}

.mode-selector button.active {
  background: var(--primary-color);
  color: white;
  font-weight: bold;
}

.control-panel {
  background: var(--card-bg);
  border-radius: 8px;
  padding: 20px;
  box-shadow: 0 4px 12px rgba(0,0,0,0.08);
  margin-bottom: 20px;
}

.form-group {
  margin-bottom: 15px;
}

.form-group label {
  display: block;
  margin-bottom: 5px;
  font-weight: 500;
}

.form-group input, .form-group select {
  width: 100%;
  padding: 10px;
  border: 1px solid var(--border-color);
  border-radius: 4px;
  background: var(--card-bg);
  color: var(--text-color);
}

.file-upload {
  margin-bottom: 20px;
}

.file-preview {
  display: flex;
  align-items: center;
  margin-top: 10px;
  padding: 10px;
  background: rgba(52, 152, 219, 0.1);
  border-radius: 4px;
}

.file-preview span {
  flex: 1;
}

.file-preview button {
  background: #e74c3c;
  color: white;
  border: none;
  padding: 5px 10px;
  border-radius: 4px;
  cursor: pointer;
}

.action-buttons {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
  gap: 10px;
  margin: 20px 0;
}

.action-buttons button {
  padding: 12px;
  background: var(--primary-color);
  color: white;
  border: none;
  border-radius: 4px;
  cursor: pointer;
  font-size: 16px;
  transition: background 0.3s;
}

.action-buttons button:hover {
  background: #2980b9;
}

.action-buttons button:disabled {
  background: #95a5a6;
  cursor: not-allowed;
}

.progress-section {
  margin-top: 20px;
}

.progress-bar {
  height: 20px;
  background: var(--progress-bg);
  border-radius: 10px;
  overflow: hidden;
}

.progress-fill {
  height: 100%;
  background: var(--progress-fill);
  border-radius: 10px;
  transition: width 0.3s;
}

.progress-label {
  text-align: center;
  margin-top: 5px;
  font-size: 14px;
  color: #7f8c8d;
}

.qr-preview-section {
  background: var(--card-bg);
  border-radius: 8px;
  padding: 20px;
  box-shadow: 0 4px 12px rgba(0,0,0,0.08);
  margin-bottom: 20px;
}

.qr-grid {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(200px, 1fr));
  gap: 20px;
  margin: 20px 0;
}

.qr-card {
  background: white;
  border-radius: 8px;
  overflow: hidden;
  box-shadow: 0 3px 10px rgba(0,0,0,0.1);
}

.qr-card img {
  width: 100%;
  display: block;
}

.qr-name {
  padding: 10px;
  text-align: center;
  font-size: 14px;
  background: #f8f9fa;
}

.log-section {
  background: var(--card-bg);
  border-radius: 8px;
  padding: 20px;
  box-shadow: 0 4px 12px rgba(0,0,0,0.08);
}

.log-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 15px;
}

.log-header button {
  background: #95a5a6;
  color: white;
  border: none;
  padding: 6px 12px;
  border-radius: 4px;
  cursor: pointer;
}

.log-content {
  max-height: 300px;
  overflow-y: auto;
  background: rgba(0,0,0,0.05);
  border-radius: 6px;
  padding: 15px;
}

.log-entry {
  padding: 8px 0;
  border-bottom: 1px solid var(--border-color);
  font-size: 14px;
  font-family: 'Courier New', monospace;
}

.restore-result {
  background: var(--secondary-color);
  color: white;
  border-radius: 8px;
  padding: 20px;
  text-align: center;
  margin-top: 20px;
}

.restore-result button {
  background: white;
  color: var(--secondary-color);
  border: none;
  padding: 10px 20px;
  border-radius: 4px;
  margin-top: 10px;
  cursor: pointer;
  font-weight: bold;
}
</style>