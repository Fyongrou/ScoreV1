// 全局变量，存储所有成绩数据
let allGrades = [];
// 存储当前查询结果
let currentResults = [];

// 提取错误处理函数
function handleError(error, message) {
    console.error(message, error);
    alert(message);
}

// 初始化事件监听器
function initEventListeners() {
    try {
        const elements = {
            downloadTemplateBtn: document.getElementById('downloadTemplate'),
            importDataBtn: document.getElementById('importData'),
            uploadFileInput: document.getElementById('uploadFile'),
            searchButton: document.getElementById('searchButton'),
            exportResultBtn: document.getElementById('exportResult')
        };

        for (const [key, element] of Object.entries(elements)) {
            if (!element) {
                console.error(`未找到 DOM 元素: ${key.replace('Btn', '')}`);
                return;
            }
        }

        elements.downloadTemplateBtn.addEventListener('click', downloadTemplate);
        elements.importDataBtn.addEventListener('click', () => {
            elements.uploadFileInput.click();
        });
        elements.uploadFileInput.addEventListener('change', handleFileUpload);
        elements.searchButton.addEventListener('click', performSearch);
        elements.exportResultBtn.addEventListener('click', exportResultsToExcel);
    } catch (error) {
        handleError(error, '事件监听器初始化出错');
    }
}

// 下载 Excel 模板
function downloadTemplate() {
    try {
        const templateData = [
            ['学年', '考试类型', '年级', '学号', '姓名', '班级', '语文', '数学', '英语', '物理', '化学', '道德与法治', '历史', '地理', '生物', '体育', '音乐', '美术', '信息']
        ];
        const ws = XLSX.utils.aoa_to_sheet(templateData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
        XLSX.writeFile(wb, '成绩模板.xlsx');
    } catch (error) {
        handleError(error, '下载模板出错，请检查控制台信息');
    }
}

// 填充年度下拉列表
function populateYearDropdowns() {
    try {
        const years = new Set(allGrades.map(grade => String(grade['学年'])));
        const sortedYears = Array.from(years).sort();

        const startYearSelect = document.getElementById('startYear');
        const endYearSelect = document.getElementById('endYear');

        // 清空现有选项
        startYearSelect.innerHTML = '<option value="">起始年度</option>';
        endYearSelect.innerHTML = '<option value="">结束年度</option>';

        sortedYears.forEach(year => {
            const startOption = document.createElement('option');
            startOption.value = year;
            startOption.textContent = year;
            startYearSelect.appendChild(startOption);

            const endOption = document.createElement('option');
            endOption.value = year;
            endOption.textContent = year;
            endYearSelect.appendChild(endOption);
        });
    } catch (error) {
        handleError(error, '填充年度下拉列表出错');
    }
}

// 数据校验
function validateData(data) {
    const requiredColumns = [
        '学年', '考试类型', '年级', '学号', '姓名', '班级', '语文', '数学', '英语', '物理', '化学', '道德与法治', '历史', '地理', '生物', '体育', '音乐', '美术', '信息'
    ];
    const gradeColumns = requiredColumns.slice(6);
    const validGrades = [];
    const errorRecords = [];

    data.forEach((row, rowIndex) => {
        let hasError = false;
        const errorInfo = { ...row, 错误信息: '' };

        // 检查是否包含所有必要列
        for (let col of requiredColumns) {
            if (!(col in row)) {
                errorInfo.错误信息 += `缺少列: ${col}; `;
                hasError = true;
            }
        }

        // 检查考试类型
        if (!['期中', '期末'].includes(row['考试类型'])) {
            errorInfo.错误信息 += `无效的考试类型: ${row['考试类型']}; `;
            hasError = true;
        }

        // 检查成绩是否为数值或预设等级
        for (let col of gradeColumns) {
            const grade = row[col];
            if (isNaN(grade) && !['A', 'B', 'C', 'D'].includes(grade)) {
                errorInfo.错误信息 += `无效的成绩值: ${grade} 在 ${col}; `;
                hasError = true;
            }
        }

        if (hasError) {
            errorRecords.push(errorInfo);
        } else {
            validGrades.push(row);
        }
    });

    return validGrades.length === data.length ? { validGrades, errorRecords: [] } : { validGrades: null, errorRecords };
}

// 导出错误报告
function exportErrorReport(errorRecords) {
    try {
        const ws = XLSX.utils.json_to_sheet(errorRecords);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, '错误报告');
        XLSX.writeFile(wb, '数据导入错误报告.xlsx');
    } catch (error) {
        handleError(error, '导出错误报告出错，请检查控制台信息');
    }
}

// 执行查询
async function performSearch() {
    try {
        console.log('开始执行查询');
        const searchId = document.getElementById('searchId')?.value.trim() || '';
        const searchName = document.getElementById('searchName')?.value.trim() || '';
        const startYear = document.getElementById('startYear')?.value;
        const endYear = document.getElementById('endYear')?.value;
        const searchType = document.getElementById('searchType')?.value;

        const url = `http://localhost:3000/search?searchId=${searchId}&searchName=${searchName}&startYear=${startYear}&endYear=${endYear}&searchType=${searchType}`;
        console.log('请求的 URL:', url);

        const response = await fetch(url);

        if (!response.ok) {
            throw new Error(`HTTP 错误! 状态码: ${response.status}`);
        }

        const data = await response.json();

        if (data.success) {
            currentResults = data.results;
            renderResults(currentResults);
        } else {
            alert('查询失败: ' + data.message);
        }
    } catch (error) {
        console.error('查询出错的详细信息:', error);
        if (error.message.includes('Failed to fetch')) {
            handleError(error, '查询出错，请检查网络连接或服务器是否启动');
        } else {
            handleError(error, '查询出错，请检查控制台信息');
        }
    }
}

// 渲染查询结果
function renderResults(results) {
    try {
        const tableBody = document.getElementById('resultTable')?.getElementsByTagName('tbody')[0];
        if (!tableBody) {
            console.error('未找到结果表格的 tbody 元素');
            return;
        }
        tableBody.innerHTML = '';

        if (results.length === 0) {
            const row = tableBody.insertRow();
            const cell = row.insertCell(0);
            cell.colSpan = 19;
            cell.textContent = '未找到符合条件的记录';
            return;
        }

        results.forEach((grade) => {
            const row = tableBody.insertRow();
            const columns = [
                '学年', '考试类型', '年级', '学号', '姓名', '班级', '语文', '数学', '英语', '物理', '化学', '道德与法治', '历史', '地理', '生物', '体育', '音乐', '美术', '信息'
            ];

            columns.forEach((col, colIndex) => {
                const cell = row.insertCell(colIndex);
                if (colIndex >= 6) {
                    const input = document.createElement('input');
                    input.value = grade[col] || '';
                    input.classList.add('grade-input');
                    input.addEventListener('change', (e) => {
                        updateGrade(col, grade, e.target.value);
                    });
                    cell.appendChild(input);
                } else {
                    cell.textContent = grade[col] || '';
                }
            });
        });
    } catch (error) {
        handleError(error, '渲染结果出错，请检查控制台信息');
    }
}

// 更新成绩
async function updateGrade(column, grade, newValue) {
    try {
        if (isNaN(newValue) && !['A', 'B', 'C', 'D'].includes(newValue)) {
            alert('成绩值无效，请输入数值或 A、B、C、D');
            return;
        }
        // 这里需要将更新发送到服务器
        const response = await fetch('http://localhost:3000/update', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                id: grade.id, // 假设每条记录有唯一的 id
                column: column,
                newValue: newValue
            })
        });
        const data = await response.json();

        if (data.success) {
            grade[column] = newValue;
            alert('成绩更新成功');
        } else {
            alert('成绩更新失败: ' + data.message);
        }
    } catch (error) {
        handleError(error, '更新成绩出错，请检查网络连接');
    }
}

// 导出查询结果到 Excel
function exportResultsToExcel() {
    try {
        if (currentResults.length === 0) {
            alert('没有可导出的查询结果');
            return;
        }
        const ws = XLSX.utils.json_to_sheet(currentResults);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, '查询结果');
        XLSX.writeFile(wb, '查询结果.xlsx');
    } catch (error) {
        handleError(error, '导出结果到 Excel 出错，请检查控制台信息');
    }
}

// 初始化
window.addEventListener('load', () => {
    try {
        initEventListeners();
        console.log('事件监听器初始化完成');
    } catch (error) {
        handleError(error, '页面加载初始化出错，请检查控制台信息');
    }
});

// 初始化 IndexedDB 数据库
const dbName = 'DC';
const storeName = 'ScoreV1';
let db;

// 打开数据库
const request = indexedDB.open(dbName, 1);

// 当数据库版本更新或首次创建时执行
request.onupgradeneeded = function(event) {
    db = event.target.result;
    // 检查对象仓库是否存在，如果不存在则创建
    if (!db.objectStoreNames.contains(storeName)) {
        const objectStore = db.createObjectStore(storeName, {
            keyPath: 'id', // 主键字段
            autoIncrement: true // 自动递增主键
        });

        // 为常用查询字段创建索引
        objectStore.createIndex('studentId', 'studentId', { unique: false });
        objectStore.createIndex('examType', 'examType', { unique: false });
        objectStore.createIndex('schoolYear', 'schoolYear', { unique: false });
    }
};

// 数据库打开成功时执行
request.onsuccess = function(event) {
    db = event.target.result;
    console.log('数据库 DC 已成功打开');
};

// 数据库打开失败时执行
request.onerror = function(event) {
    console.error('数据库 DC 打开失败:', event.target.error);
};

// 关闭数据库连接的函数
function closeDB() {
    if (db) {
        db.close();
        console.log('数据库 DC 已关闭');
    }
}

// 示例：添加数据到数据库
function addData(data) {
    if (db) {
        const transaction = db.transaction([storeName], 'readwrite');
        const objectStore = transaction.objectStore(storeName);
        const addRequest = objectStore.add(data);

        addRequest.onsuccess = function(event) {
            console.log('数据添加成功，ID:', event.target.result);
        };

        addRequest.onerror = function(event) {
            console.error('数据添加失败:', event.target.error);
        };
    }
}

// 示例：从数据库获取数据
function getData(id) {
    if (db) {
        const transaction = db.transaction([storeName], 'readonly');
        const objectStore = transaction.objectStore(storeName);
        const getRequest = objectStore.get(id);

        getRequest.onsuccess = function(event) {
            if (event.target.result) {
                console.log('获取到的数据:', event.target.result);
            } else {
                console.log('未找到 ID 为 ' + id + ' 的数据');
            }
        };

        getRequest.onerror = function(event) {
            console.error('数据获取失败:', event.target.error);
        };
    }
}

// 处理文件上传
function handleFileUpload(event) {
    try {
        const file = event.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet);

            const { validGrades, errorRecords } = validateData(jsonData);
            if (errorRecords.length > 0) {
                exportErrorReport(errorRecords);
                alert('数据导入存在错误，请查看错误报告');
                return;
            }

            allGrades = validGrades;
            populateYearDropdowns();
            alert('数据导入成功');
        };
        reader.readAsArrayBuffer(file);
    } catch (error) {
        handleError(error, '文件上传处理出错，请检查控制台信息');
    }
}