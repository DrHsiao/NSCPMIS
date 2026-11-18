// 預設單一檔案大小（MB）
const DEFAULT_MAX_FILE_SIZE_MB = 20;
// 預設所有檔案大小（MB）
const DEFAULT_TOTAL_FILE_SIZE_MB = 20;
// 預設檔案類型
const DEFAULT_ALLOWED_FILE_TYPES = ['jpg', 'jpeg', 'png', 'pdf'];
// 預設 MIME 類型
const DEFAULT_ALLOWED_MIME_TYPES = ['image/jpeg', 'image/png', 'application/pdf'];


// 檢查副檔名
function isAllowedFileType(file, allowedTypes = DEFAULT_ALLOWED_FILE_TYPES) {
    if (!file || !file.name) return false;
    const extension = file.name.split('.').pop().toLowerCase();
    return allowedTypes.includes(extension);
}

// 檢查 MIME 類型
function isAllowedMimeType(file, allowedMimes = DEFAULT_ALLOWED_MIME_TYPES) {
    if (!file || !file.type) return false;
    return allowedMimes.includes(file.type);
}

// 檢查檔案大小（單位 MB）
function isAllowedFileSize(file, maxSizeMB = DEFAULT_MAX_FILE_SIZE_MB) {
    if (!file || typeof file.size !== 'number') return false;
    const maxSize = maxSizeMB * 1024 * 1024;
    return file.size <= maxSize;
}

// 檢查所有檔案的總大小（單位 MB）
function validateAllFiles(maxTotalMB = DEFAULT_TOTAL_FILE_SIZE_MB) {
    const inputs = document.querySelectorAll("input[type='file']");
    let totalSize = 0;

    inputs.forEach(input => {
        for (const file of input.files) {
            totalSize += file.size;
        }
    });

    const maxTotalSize = maxTotalMB * 1024 * 1024;

    return totalSize <= maxTotalSize;
}


// 驗證單一檔案
function validateSingleFile(file, options = {}) {
    const {
        allowedTypes = DEFAULT_ALLOWED_FILE_TYPES,
        allowedMimes = DEFAULT_ALLOWED_MIME_TYPES,
        maxSizeMB = DEFAULT_MAX_FILE_SIZE_MB,
        maxTotalMB = DEFAULT_TOTAL_FILE_SIZE_MB,
        checkMime = true
    } = options;

    if (!isAllowedFileType(file, allowedTypes)) {
        return { valid: false, reason: '檔案格式不正確（允許類型：' + allowedTypes.join(', ') + '）' };
    }

    if (checkMime && !isAllowedMimeType(file, allowedMimes)) {
        return { valid: false, reason: '檔案 MIME 類型不被允許' };
    }

    if (!isAllowedFileSize(file, maxSizeMB)) {
        return { valid: false, reason: '檔案大小不能超過 ' + maxSizeMB + 'MB' };
    }

    if (!validateAllFiles(maxTotalMB)) {
        return { valid: false, reason: '上傳的檔案總大小不可超過 ' + DEFAULT_TOTAL_FILE_SIZE_MB + 'MB' };
    }

    return { valid: true };
}