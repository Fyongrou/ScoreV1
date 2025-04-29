// 全局变量，存储所有成绩数据
let allGrades = [];
// 存储当前查询结果
let currentResults = [];

// 初始化事件监听器
function initEventListeners() {
    try {
        const downloadTemplateBtn = document.getElementById('downloadTemplate');
        const importDataBtn = document.getElementById('importData');
        const uploadFileInput = document.getElementById('uploadFile');
        const searchButton = document.getElementById('searchButton');
        const exportResultBtn = document.getElementById('exportResult');

        if (!downloadTemplateBtn || !importDataBtn || !uploadFileInput || !searchButton || !exportResultBtn) {
            console.error('部分 DOM 元素未找到，请检查 HTML 结构');
            return;
        }

        downloadTemplateBtn.addEventListener('click', downloadTemplate);
        importDataBtn.addEventListener('click', () => {
            uploadFileInput.click();
        });
        uploadFileInput.addEventListener('change', handleFileUpload);
        searchButton.addEventListener('click', performSearch);
        exportResultBtn.addEventListener('click', exportResultsToExcel);
    } catch (error) {
        console.error('事件监听器初始化出错:', error);
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
        console.error('下载模板出错:', error);
        alert('下载模板出错，请检查控制台信息');
    }
}

// 处理文件上传
async function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    const formData = new FormData();
    formData.append('file', file);

    try {
        const response = await fetch('http://localhost:3000/upload', {
            method: 'POST',
            body: formData
        });

        if (response.ok) {
            const data = await response.json();
            if (data.success) {
                alert('数据上传成功，已保存到服务器');
                // 可以在这里重新获取最新数据
            } else {
                alert('数据上传失败: ' + data.message);
            }
        } else {
            alert('服务器响应错误');
        }
    } catch (error) {
        console.error('文件上传出错:', error);
        alert('文件上传出错，请检查网络连接');
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
        console.error('填充年度下拉列表出错:', error);
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
        console.error('导出错误报告出错:', error);
        alert('导出错误报告出错，请检查控制台信息');
    }
}

// 执行查询
function performSearch() {
    try {
        console.log('开始执行查询');
        const searchId = document.getElementById('searchId')?.value.trim() || '';
        const searchName = document.getElementById('searchName')?.value.trim() || '';
        const startYear = document.getElementById('startYear')?.value;
        const endYear = document.getElementById('endYear')?.value;
        const searchType = document.getElementById('searchType')?.value;

        // 这里需要从服务器获取数据
        fetch(`http://localhost:3000/search?searchId=${searchId}&searchName=${searchName}&startYear=${startYear}&endYear=${endYear}&searchType=${searchType}`)
          .then(response => response.json())
          .then(data => {
                if (data.success) {
                    currentResults = data.results;
                    renderResults(currentResults);
                } else {
                    alert('查询失败: ' + data.message);
                }
            })
          .catch(error => {
                console.error('查询出错:', error);
                alert('查询出错，请检查网络连接');
            });
    } catch (error) {
        console.error('执行查询出错:', error);
        alert('执行查询出错，请检查控制台信息');
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
        console.error('渲染结果出错:', error);
        alert('渲染结果出错，请检查控制台信息');
    }
}

// 更新成绩
function updateGrade(column, grade, newValue) {
    try {
        if (isNaN(newValue) && !['A', 'B', 'C', 'D'].includes(newValue)) {
            alert('成绩值无效，请输入数值或 A、B、C、D');
            return;
        }
        // 这里需要将更新发送到服务器
        fetch('http://localhost:3000/update', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                id: grade.id, // 假设每条记录有唯一的 id
                column: column,
                newValue: newValue
            })
        })
          .then(response => response.json())
          .then(data => {
                if (data.success) {
                    grade[column] = newValue;
                    alert('成绩更新成功');
                } else {
                    alert('成绩更新失败: ' + data.message);
                }
            })
          .catch(error => {
                console.error('更新成绩出错:', error);
                alert('更新成绩出错，请检查网络连接');
            });
    } catch (error) {
        console.error('更新成绩出错:', error);
        alert('更新成绩出错，请检查控制台信息');
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
        console.error('导出结果到 Excel 出错:', error);
        alert('导出结果到 Excel 出错，请检查控制台信息');
    }
}

// 初始化
window.addEventListener('load', () => {
    try {
        initEventListeners();
        console.log('事件监听器初始化完成');
    } catch (error) {
        console.error('页面加载初始化出错:', error);
        alert('页面加载初始化出错，请检查控制台信息');
    }
});