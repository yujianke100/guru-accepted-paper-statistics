document.getElementById('addRowButton').addEventListener('click', () => {
    const inputContainer = document.getElementById('inputContainer');
    const newRow = document.createElement('div');
    newRow.className = 'inputRow';
    newRow.innerHTML = `
        <input type="text" class="remarkInput" placeholder="备注">
        <textarea class="urlInput" placeholder="网址"></textarea>
        <input type="file" class="fileInput" accept=".xlsx">
        <button class="deleteRowButton">-</button>
    `;
    inputContainer.appendChild(newRow);

    const fileInput = newRow.querySelector('.fileInput');
    fileInput.addEventListener('change', handleFileUpload);

    const deleteButton = newRow.querySelector('.deleteRowButton');
    deleteButton.addEventListener('click', () => {
        inputContainer.removeChild(newRow);
    });

    const urlInput = newRow.querySelector('.urlInput');
    urlInput.addEventListener('input', adjustHeight); // 添加动态调整高度事件
});

document.querySelectorAll('.urlInput').forEach(urlInput => {
    urlInput.addEventListener('input', adjustHeight); // 确保现有输入框也绑定动态调整高度事件
});

function adjustHeight(event) {
    const textarea = event.target;
    textarea.style.height = 'auto'; // 先重置高度
    textarea.style.height = `${textarea.scrollHeight}px`; // 根据内容调整高度
}

async function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    const row = event.target.closest('.inputRow');
    const urlInput = row.querySelector('.urlInput');
    urlInput.disabled = true; // 禁用网址输入框

    const reader = new FileReader();
    reader.onload = async (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        const authorColumn = json.find(row => row.includes('Authors'));
        if (!authorColumn) {
            console.error('未找到 "Authors" 列');
            return;
        }

        const authorIndex = authorColumn.indexOf('Authors');
        const authors = json.slice(1).map(row => row[authorIndex]).filter(Boolean);

        const authorCounts = {};
        authors.forEach(authorLine => {
            const authorNames = authorLine.split(';').map(name => name.split('(')[0].trim());
            authorNames.forEach(name => {
                authorCounts[name] = (authorCounts[name] || 0) + 1;
            });
        });

        console.log('作者统计结果:', authorCounts);
    };

    reader.readAsArrayBuffer(file);
}

document.getElementById('processButton').addEventListener('click', async () => {
    const inputRows = document.querySelectorAll('.inputRow');
    const authorStats = {};
    const remarkMap = {}; // 用于记录备注与其对应的链接索引
    const remarks = [];
    const allAuthors = []; // 用于存储所有读取到的作者

    console.log('开始处理输入行...'); // 添加调试日志

    for (const row of inputRows) {
        const remark = row.querySelector('.remarkInput').value.trim() || 'No Remark'; // 默认备注为 "No Remark"
        const urls = row.querySelector('.urlInput').value.trim().split('\n').filter(url => url.trim());
        const fileInput = row.querySelector('.fileInput');
        const file = fileInput.files[0];

        console.log(`处理行: 备注=${remark}, 文件=${file ? file.name : '无文件'}, URL数量=${urls.length}`); // 添加调试日志

        if (!remarkMap[remark]) {
            remarkMap[remark] = remarks.length;
            remarks.push(remark);
        }
        const remarkIndex = remarkMap[remark];

        // 处理网址
        for (const url of urls) {
            try {
                const response = await fetch(url.trim());
                const data = await response.json();
                const hits = data.result.hits.hit || [];
                hits.forEach(hit => {
                    let authors = hit.info.authors?.author || []; // 使用 `let` 以允许重新赋值
                    if (!Array.isArray(authors)) {
                        // 如果 authors 不是数组，将其转换为数组
                        authors = [authors];
                    }
                    authors.forEach(author => {
                        if (!author || !author.text) {
                            console.warn(`URL: ${url.trim()} 的作者信息无效`, author);
                            return;
                        }
                        const name = author.text.trim();
                        if (!authorStats[name]) {
                            authorStats[name] = Array(remarks.length).fill(0);
                        }
                        // 确保数组长度与备注数量一致
                        while (authorStats[name].length < remarks.length) {
                            authorStats[name].push(0);
                        }
                        authorStats[name][remarkIndex]++;
                    });
                });
            } catch (error) {
                console.error(`无法处理 URL: ${url.trim()}`, error);
            }
        }

        // 处理上传的 Excel 文件
        if (file) {
            const reader = new FileReader();
            reader.onload = (e) => {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const sheetName = workbook.SheetNames[0];
                const sheet = workbook.Sheets[sheetName];
                const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });

                console.log('读取到的 Excel 数据:', json.slice(0, 5)); // 打印前 5 行 Excel 数据

                if (!json || json.length === 0) {
                    console.error('Excel 文件为空或无法解析');
                    return;
                }

                const authorColumnIndex = json[0].findIndex(header => header.includes('Authors'));
                if (authorColumnIndex === -1) {
                    console.error('未找到 "Authors" 列');
                    return;
                }

                json.slice(1).forEach(row => {
                    const authorLine = row[authorColumnIndex];
                    if (authorLine) {
                        // 支持分号或逗号分隔的作者列表
                        const authorNames = authorLine.split(/;|,/).map(name => name.split('(')[0].trim());
                        authorNames.forEach(name => {
                            if (!name) return; // 跳过空作者
                            allAuthors.push(name); // 将作者添加到列表中
                            if (!authorStats[name]) {
                                authorStats[name] = Array(remarks.length).fill(0);
                            }
                            // 确保数组长度与备注数量一致
                            while (authorStats[name].length < remarks.length) {
                                authorStats[name].push(0);
                            }
                            authorStats[name][remarkIndex]++;
                        });
                    }
                });

                // 打印前 5 个作者到命令行
                console.log('读取到的作者列表前 5 行:', allAuthors.slice(0, 5));
            };
            await new Promise(resolve => {
                reader.onloadend = resolve;
                reader.readAsArrayBuffer(file);
            });
        }
    }

    // 导出 Excel 表格
    const rows = [['Author Name', ...remarks, 'Total']]; // 确保表头完整
    Object.keys(authorStats).forEach(author => {
        // 确保每列都有数据，填充缺失的列为 0
        while (authorStats[author].length < remarks.length) {
            authorStats[author].push(0);
        }
        rows.push([author, ...authorStats[author], '']); // Total 列留空
    });

    if (rows.length === 1) {
        console.error('没有数据可导出');
        return;
    }

    console.log('生成的统计结果:', rows.slice(0, 5)); // 打印前 5 行统计结果
    exportToExcel(rows);
});

function exportToExcel(rows) {
    const csvContent = rows.map(row => row.map(cell => `"${cell}"`).join(',')).join('\n'); // 确保标题和内容正确编码
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = '作者统计结果.csv';
    link.click();
}
