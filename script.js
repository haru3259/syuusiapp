document.addEventListener('DOMContentLoaded', function () {
    // --- DOM Elements ---
    const appContainer = document.getElementById('app-container');
    const editModeToggle = document.getElementById('edit-mode-toggle');
    const summaryButton = document.getElementById('summary-button');
    const resetDataButton = document.getElementById('reset-data-button');
    const searchBox = document.getElementById('search-box');
    const exportJsonButton = document.getElementById('export-json-button');
    const importJsonButton = document.getElementById('import-json-button');
    const bulkEditButton = document.getElementById('bulk-edit-button');
    const settingsButton = document.getElementById('settings-button');

    const userNameDisplay = document.getElementById('user-name-display');
    const userNameEdit = document.getElementById('user-name-edit');

    // Password Modal
    const passwordModal = document.getElementById('password-modal');
    const passwordForm = document.getElementById('password-form');
    const passwordInput = document.getElementById('password-input');
    const passwordError = document.getElementById('password-error');

    // Edit Modal
    const modal = document.getElementById('modal');
    const modalForm = document.getElementById('modal-form');
    const modalTitle = document.getElementById('modal-title');
    const itemNameInput = document.getElementById('item-name');
    const itemAmountInput = document.getElementById('item-amount');
    const cancelButton = document.getElementById('cancel-button');

    // Bulk Edit Modal
    const bulkEditModal = document.getElementById('bulk-edit-modal');
    const bulkEditForm = document.getElementById('bulk-edit-form');
    const bulkEditSelect = document.getElementById('bulk-edit-select');
    const bulkEditNewName = document.getElementById('bulk-edit-new-name');
    const bulkEditNewAmount = document.getElementById('bulk-edit-new-amount');
    const bulkCancelButton = document.getElementById('bulk-cancel-button');

    // Summary Modal
    const summaryModal = document.getElementById('summary-modal');
    const summaryCloseButton = document.getElementById('summary-close-button');
    const summaryContent = document.getElementById('summary-content');

    // Category Summary Modal
    const categorySummaryModal = document.getElementById('category-summary-modal');
    const categorySummaryButton = document.getElementById('category-summary-button');
    const categorySummaryCloseButton = document.getElementById('category-summary-close-button');
    const categorySummaryContent = document.getElementById('category-summary-content');

    // New Data Modal
    const newDataButton = document.getElementById('new-data-button');
    const newDataModal = document.getElementById('new-data-modal');
    const newDataForm = document.getElementById('new-data-form');
    const newDataCancelButton = document.getElementById('new-data-cancel-button');
    const newDataCloseButton = document.getElementById('new-data-close-button');
    const startYearInput = document.getElementById('start-year');
    const initialBalanceInput = document.getElementById('initial-balance');
    const userNameInput = document.getElementById('user-name');

    // Add Year Modal
    const addYearModal = document.getElementById('add-year-modal');
    const addYearForm = document.getElementById('add-year-form');
    const addYearCloseButton = document.getElementById('add-year-close-button');
    const addYearCancelButton = document.getElementById('add-year-cancel-button');
    const newYearNameInput = document.getElementById('new-year-name');
    const addYearMonthsContainer = document.getElementById('add-year-months-container');

    // Edit Months Modal
    const editMonthsModal = document.getElementById('edit-months-modal');
    const editMonthsForm = document.getElementById('edit-months-form');
    const editMonthsTitle = document.getElementById('edit-months-title');
    const editMonthsContainer = document.getElementById('edit-months-container');
    const editMonthsCloseButton = document.getElementById('edit-months-close-button');
    const editMonthsCancelButton = document.getElementById('edit-months-cancel-button');
    const deleteYearButton = document.getElementById('delete-year-button');

    // Category Settings Modal (now part of general settings modal)
    const settingsModal = document.getElementById('settings-modal');
    const settingsCloseButton = document.getElementById('settings-close-button');
    const categoryInput = document.getElementById('category-input');
    const saveSettingsButton = document.getElementById('save-settings');
    const cancelCategoriesButton = document.getElementById('cancel-categories-button');
    const autoGetCategoriesButton = document.getElementById('auto-get-categories-button');
    const passwordSettingInput = document.getElementById('password-setting');

    // --- App State ---
    let state = {
        isEditMode: false,
        financialData: [],
        userName: '',
        editingItem: null, // { yearId, monthId, type, itemId }
        searchTerm: '',
        categories: [],
        password: ''
    };

    // --- SVG Icons ---
    const icons = {
        add: '<svg viewBox="0 0 24 24"><path d="M19 13h-6v6h-2v-6H5v-2h6V5h2v6h6v2z"/></svg>',
        edit: '<svg viewBox="0 0 24 24"><path d="M3 17.25V21h3.75L17.81 9.94l-3.75-3.75L3 17.25zM20.71 7.04c.39-.39.39-1.02 0-1.41l-2.34-2.34c-.39-.39-1.02-.39-1.41 0l-1.83 1.83 3.75 3.75 1.83-1.83z"/></svg>',
        delete: '<svg viewBox="0 0 24 24"><path d="M6 19c0 1.1.9 2 2 2h8c1.1 0 2-.9 2-2V7H6v12zM19 4h-3.5l-1-1h-5l-1 1H5v2h14V4z"/></svg>',
    };

    // --- Data Initialization ---
    async function initializeApp() {
        const storedPassword = localStorage.getItem('password');
        if (storedPassword && storedPassword.length > 0) {
            state.password = storedPassword;
            passwordModal.classList.add('visible');
            passwordInput.focus();
        } else {
            // No password, load data immediately
            await loadAndRenderData();
        }
    }

    async function loadAndRenderData(enteredPassword = null) {
        const storedData = localStorage.getItem('financialData');
        const storedName = localStorage.getItem('userName');

        if (storedName) {
            state.userName = storedName;
            userNameDisplay.textContent = `${state.userName}の収支`;
        }

        state.categories = loadCategories();

        if (storedData) {
            try {
                let parsedData;
                const passwordToUse = enteredPassword || state.password;

                if (passwordToUse) {
                    const bytes = CryptoJS.AES.decrypt(storedData, passwordToUse);
                    const decryptedString = bytes.toString(CryptoJS.enc.Utf8);
                    if (!decryptedString) {
                        // This error will be caught and handled below
                        throw new Error("Decryption failed. Invalid password?");
                    }
                    parsedData = JSON.parse(decryptedString);
                } else {
                    parsedData = JSON.parse(storedData);
                }
                state.financialData = processData(parsedData);
            } catch (err) {
                console.error('データの読み込みに失敗しました:', err);
                // If an entered password fails, we want to show the error on the password modal
                if (enteredPassword) {
                    passwordError.textContent = 'パスワードが違います。';
                    passwordInput.value = '';
                    passwordInput.focus();
                    return; // Stop execution
                }

                // Fallback for non-password related errors or old key
                try {
                    const bytes = CryptoJS.AES.decrypt(storedData, "YourSecretKey");
                    const decryptedString = bytes.toString(CryptoJS.enc.Utf8);
                    if (!decryptedString) throw new Error("Old key decryption failed");
                    const parsedData = JSON.parse(decryptedString);
                    state.financialData = processData(parsedData);
                    showPopupMessage('古い暗号化データを検出しました。設定からパスワードを再設定・保存してください。');
                } catch (fallbackErr) {
                    console.error('フォールバック復号にも失敗しました:', fallbackErr);
                    if (confirm('データの読み込みに失敗しました。データをリセットしてもよろしいですか？')) {
                        state.financialData = await loadDefaultData();
                    } else {
                        state.financialData = [];
                    }
                }
            }
        } else {
            state.financialData = await loadDefaultData();
        }

        // If we got here, decryption was successful, so hide the password modal
        if (passwordModal.classList.contains('visible')) {
            passwordModal.style.opacity = '0';
            setTimeout(() => {
                passwordModal.classList.remove('visible');
                passwordModal.style.opacity = '1';
            }, 500);
        }
        rerender();
    }

    async function loadDefaultData() {
        try {
            const response = await fetch('default-data.json');
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            const defaultData = await response.json();
            return processData(defaultData);
        } catch (error) {
            console.error('デフォルトデータの読み込みに失敗しました:', error);
            showPopupMessage('アプリケーションの初期化に失敗しました。');
            return [];
        }
    }

    function processData(data) {
        return data.map(year => ({
            ...year,
            id: year.id || `year_${Date.now()}_${Math.random()}`,
            months: year.months.map(month => ({
                ...month,
                id: month.id || `month_${Date.now()}_${Math.random()}`,
                note: month.note || '',
                startingBalance: Number(month.startingBalance) || 0,
                income: (month.income || []).map(item => ({
                    ...item,
                    id: item.id || `item_${Date.now()}_${Math.random()}`,
                    amount: Number(item.amount) || 0
                })),
                expenditure: (month.expenditure || []).map(item => ({
                    ...item,
                    id: item.id || `item_${Date.now()}_${Math.random()}`,
                    amount: Number(item.amount) || 0
                }))
            }))
        }));
    }

    // --- State Management & Calculations ---
    const saveState = () => {
        let dataToStore = JSON.stringify(state.financialData);
        if (state.password) {
            dataToStore = CryptoJS.AES.encrypt(dataToStore, state.password).toString();
        }
        localStorage.setItem('financialData', dataToStore);
        localStorage.setItem('userName', state.userName);
        localStorage.setItem('password', state.password);
        saveCategories(state.categories);
    };

    const saveCategories = (categories) => {
        localStorage.setItem('categories', JSON.stringify(categories));
    };

    const loadCategories = () => {
        const storedCategories = localStorage.getItem('categories');
        if (storedCategories) {
            return JSON.parse(storedCategories);
        } else {
            // Return the default list if nothing is in storage
            return [
                "バイト給料", "給料", "日本学生支援機構給付金", "川崎市大学等進学奨学金",
                "千文基金", "篠原財団", "家賃", "初期費用", "住宅設定費", "家電・家具",
                "光熱費", "電気ガス水道費", "インターネット", "NHK通信", "NHK通信料",
                "学費", "学費減免", "教材費", "大学進学等自立生活支度費",
                "専門学校入学支度金", "社会的自立支援費（私立）", "定期", "バス 定期",
                "電車 定期", "交通費定期", "食費", "生活消耗品", "国民健康保険",
                "娯楽費", "その他"
            ].sort();
        }
    };

    const calculateAllBalances = () => {
        let previousMonthFinalBalance = 0;
        let grandTotalIncome = 0;
        let grandTotalExpenditure = 0;
        let isFirstMonthEver = true;
        let initialBalance = 0;

        state.financialData.forEach(year => {
            year.months.sort((a, b) => {
                const monthA = parseInt(a.month);
                const monthB = parseInt(b.month);

                if (year.year.endsWith('年度') || year.year === '1年次' || year.year === '2年次') {
                    let adjustedMonthA = monthA < 4 ? monthA + 12 : monthA;
                    let adjustedMonthB = monthB < 4 ? monthB + 12 : monthB;

                    if (year.year === '1年次') {
                        if (monthA === 3 && a.note.includes('貯金')) adjustedMonthA = 3;
                        if (monthB === 3 && b.note.includes('貯金')) adjustedMonthB = 3;
                    }

                    return adjustedMonthA - adjustedMonthB;
                }
                
                return monthA - monthB;
            });

            let yearTotalIncome = 0;
            let yearTotalExpenditure = 0;

            year.months.forEach(month => {
                month.totalIncome = month.income.reduce((sum, item) => sum + (Number(item.amount) || 0), 0);
                month.totalExpenditure = month.expenditure.reduce((sum, item) => sum + (Number(item.amount) || 0), 0);

                if (isFirstMonthEver) {
                    initialBalance = month.startingBalance;
                    previousMonthFinalBalance = month.startingBalance; // 最初の月のstartingBalanceで初期化
                    isFirstMonthEver = false;
                } else {
                    month.startingBalance = previousMonthFinalBalance;
                }
                
                month.finalBalance = month.startingBalance + month.totalIncome - month.totalExpenditure;
                previousMonthFinalBalance = month.finalBalance;

                yearTotalIncome += month.totalIncome;
                yearTotalExpenditure += month.totalExpenditure;
            });

            year.totalIncome = yearTotalIncome;
            year.totalExpenditure = yearTotalExpenditure;
            if (year.months.length > 0) {
                year.finalBalance = year.months[year.months.length - 1].finalBalance;
            }
            
            grandTotalIncome += yearTotalIncome;
            grandTotalExpenditure += yearTotalExpenditure;
        });

        state.totals = {
            income: grandTotalIncome,
            expenditure: grandTotalExpenditure,
            balance: previousMonthFinalBalance,
            initial: initialBalance
        };
    };

    // --- Modals ---
    const showModal = (modalElement) => {
        document.body.classList.add('modal-open');
        modalElement.style.display = 'flex';
    };
    const hideModal = (modalElement) => {
        document.body.classList.remove('modal-open');
        modalElement.style.display = 'none';
    };

    const showEditModal = (title, name = '', amount = '') => {
        modalTitle.textContent = title;
        itemNameInput.value = name;
        itemAmountInput.value = amount;

        // Populate datalist with existing item names
        const itemNamesDatalist = document.getElementById('item-names');
        itemNamesDatalist.innerHTML = '';
        state.categories.forEach(name => {
            const option = document.createElement('option');
            option.value = name;
            itemNamesDatalist.appendChild(option);
        });

        showModal(modal);
    };

    // --- CRUD & Data Functions ---
    const handleAddItem = (yearId, monthId, type) => {
        state.editingItem = { yearId, monthId, type };
        showEditModal(`新しい${type === 'income' ? '収入' : '支出'}を追加`, '', 0);
    };

    const handleAddYear = (options) => {
        const { name, type, selectedMonths, afterYearId } = options;

        const lastMonthOfPreviousYear = findYear(afterYearId)?.months.slice(-1)[0];
        const startingBalance = lastMonthOfPreviousYear ? lastMonthOfPreviousYear.finalBalance : 0;

        const newYear = {
            id: `year_${Date.now()}`,
            year: name,
            months: selectedMonths.map((monthName, index) => ({
                id: `month_${Date.now()}_${index}`,
                month: monthName,
                note: '',
                startingBalance: index === 0 ? startingBalance : 0,
                income: [],
                expenditure: []
            }))
        };

        const afterIndex = state.financialData.findIndex(y => y.id === afterYearId);
        if (afterIndex > -1) {
            state.financialData.splice(afterIndex + 1, 0, newYear);
        } else {
            state.financialData.push(newYear);
        }

        rerender();
    };

    const handleEditItem = (yearId, monthId, type, itemId) => {
        const item = findItem(yearId, monthId, type, itemId);
        state.editingItem = { yearId, monthId, type, itemId };
        showEditModal(`${type === 'income' ? '収入' : '支出'}を編集`, item.name, item.amount || 0);
    };

    const handleDeleteItem = (yearId, monthId, type, itemId) => {
        if (!confirm('この項目を削除しますか？')) return;
        const month = findMonth(yearId, monthId);
        month[type] = month[type].filter(i => i.id !== itemId);
        rerender();
    };

    const handleYearNameUpdate = (yearId, newName) => {
        const year = findYear(yearId);
        if (year) {
            year.year = newName;
            saveState();
        }
    };

    const handleNoteUpdate = (yearId, monthId, newNote) => {
        const month = findMonth(yearId, monthId);
        month.note = newNote;
        saveState(); // No need to rerender, just save
    };

    const findYear = (yearId) => state.financialData.find(y => y.id === yearId);
    const findMonth = (yearId, monthId) => findYear(yearId)?.months.find(m => m.id === monthId);
    const findItem = (yearId, monthId, type, itemId) => findMonth(yearId, monthId)?.[type].find(i => i.id === itemId);

    const extractCategoriesFromFinancialData = () => {
        const allItems = new Set();
        state.financialData.forEach(year => {
            year.months.forEach(month => {
                month.income.forEach(item => allItems.add(item.name));
                month.expenditure.forEach(item => allItems.add(item.name));
            });
        });
        return Array.from(allItems).sort();
    };

    // --- New Data Creation ---
    const createNewFinancialData = (options) => {
        const { type, yearName, initialBalance } = options;
        const newYear = {
            id: `year_${Date.now()}`,
            year: yearName,
            months: []
        };

        let months = [];
        if (type === 'fiscal') {
            months = ['4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月', '1月', '2月', '3月'];
        } else { // calendar
            months = ['1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月'];
        }

        newYear.months = months.map((monthName, index) => {
            return {
                id: `month_${Date.now()}_${index}`,
                month: monthName,
                note: '',
                startingBalance: index === 0 ? Number(initialBalance) : 0,
                income: [],
                expenditure: []
            };
        });

        return [newYear];
    };

    // --- Summary & Export/Import ---
    const generateSummaryData = () => {
        const summary = {};
        state.financialData.forEach(year => {
            const yearSummary = { income: {}, expenditure: {} };
            year.months.forEach(month => {
                month.income.forEach(item => {
                    if (item.amount > 0) yearSummary.income[item.name] = (yearSummary.income[item.name] || 0) + item.amount;
                });
                month.expenditure.forEach(item => {
                    if (item.amount > 0) yearSummary.expenditure[item.name] = (yearSummary.expenditure[item.name] || 0) + item.amount;
                });
            });
            summary[year.year] = yearSummary;
        });
        return summary;
    };

    const renderSummaryModal = () => {
        const summaryData = generateSummaryData();
        const formatter = new Intl.NumberFormat('ja-JP', { style: 'currency', currency: 'JPY' });
        let content = '';

        for (const yearName in summaryData) {
            const yearSummary = summaryData[yearName];
            const incomeEntries = Object.entries(yearSummary.income).sort((a, b) => b[1] - a[1]);
            const expenditureEntries = Object.entries(yearSummary.expenditure).sort((a, b) => b[1] - a[1]);
            const totalIncome = incomeEntries.reduce((sum, [, amount]) => sum + amount, 0);
            const totalExpenditure = expenditureEntries.reduce((sum, [, amount]) => sum + amount, 0);

            content += `
                <div class="summary-year-section">
                    <h3 class="summary-year-title">${yearName}</h3>
                    <div class="summary-tables">
                        <div class="income">
                            <h4>収入</h4>
                            <table class="summary-table">
                                <thead><tr><th>項目</th><th class="amount">合計</th></tr></thead>
                                <tbody>${incomeEntries.map(([name, amount]) => `<tr><td>${name}</td><td class="amount">${formatter.format(amount)}</td></tr>`).join('')}</tbody>
                                <tfoot><tr><td>合計</td><td class="amount">${formatter.format(totalIncome)}</td></tr></tfoot>
                            </table>
                        </div>
                        <div class="expenditure">
                            <h4>支出</h4>
                            <table class="summary-table">
                                <thead><tr><th>項目</th><th class="amount">合計</th></tr></thead>
                                <tbody>${expenditureEntries.map(([name, amount]) => `<tr><td>${name}</td><td class="amount">${formatter.format(amount)}</td></tr>`).join('')}</tbody>
                                <tfoot><tr><td>合計</td><td class="amount">${formatter.format(totalExpenditure)}</td></tr></tfoot>
                            </table>
                        </div>
                    </div>
                </div>`;
        }
        summaryContent.innerHTML = content;
        showModal(summaryModal);
    };

    const renderCategorySummaryModal = () => {
        const incomeTotals = {};
        const expenditureTotals = {};

        state.financialData.forEach(year => {
            year.months.forEach(month => {
                month.income.forEach(item => {
                    if (item.amount > 0) {
                        incomeTotals[item.name] = (incomeTotals[item.name] || 0) + item.amount;
                    }
                });
                month.expenditure.forEach(item => {
                    if (item.amount > 0) {
                        expenditureTotals[item.name] = (expenditureTotals[item.name] || 0) + item.amount;
                    }
                });
            });
        });

        const formatter = new Intl.NumberFormat('ja-JP', { style: 'currency', currency: 'JPY' });
        const sortedIncome = Object.entries(incomeTotals).sort((a, b) => b[1] - a[1]);
        const sortedExpenditure = Object.entries(expenditureTotals).sort((a, b) => b[1] - a[1]);

        const totalIncome = sortedIncome.reduce((sum, [, amount]) => sum + amount, 0);
        const totalExpenditure = sortedExpenditure.reduce((sum, [, amount]) => sum + amount, 0);

        const incomeTable = `
            <div class="income">
                <h4>収入合計</h4>
                <table class="summary-table">
                    <thead><tr><th>項目</th><th class="amount">合計金額</th></tr></thead>
                    <tbody>${sortedIncome.map(([name, amount]) => `<tr><td>${name}</td><td class="amount">${formatter.format(amount)}</td></tr>`).join('')}</tbody>
                    <tfoot><tr><td>総合計</td><td class="amount">${formatter.format(totalIncome)}</td></tr></tfoot>
                </table>
            </div>
        `;
        const expenditureTable = `
            <div class="expenditure">
                <h4>支出合計</h4>
                <table class="summary-table">
                    <thead><tr><th>項目</th><th class="amount">合計金額</th></tr></thead>
                    <tbody>${sortedExpenditure.map(([name, amount]) => `<tr><td>${name}</td><td class="amount">${formatter.format(amount)}</td></tr>`).join('')}</tbody>
                    <tfoot><tr><td>総合計</td><td class="amount">${formatter.format(totalExpenditure)}</td></tr></tfoot>
                </table>
            </div>
        `;

        categorySummaryContent.innerHTML = `<div class="summary-tables">${incomeTable}${expenditureTable}</div>`;
        showModal(categorySummaryModal);
    };

    const ENCRYPTION_KEY = "YourSecretKey"; // 古いデータとの互換性のために残す

    const exportToJson = () => {
        const dataToExport = {
            userName: state.userName,
            financialData: state.financialData
        };

        let exportObject;
        if (state.password) {
            const encryptedData = CryptoJS.AES.encrypt(JSON.stringify(dataToExport), state.password).toString();
            exportObject = {
                isPasswordProtected: true,
                data: encryptedData
            };
            showPopupMessage('データはパスワードで暗号化されてエクスポートされました。');
        } else {
            exportObject = {
                isPasswordProtected: false,
                data: dataToExport
            };
            showPopupMessage('データは暗号化されずにエクスポートされました。');
        }
        
        const outputData = JSON.stringify(exportObject, null, 2);
        const dataBlob = new Blob([outputData], { type: "application/json" });
        const url = URL.createObjectURL(dataBlob);
        const link = document.createElement("a");
        link.setAttribute("href", url);
        link.setAttribute("download", "financial_data.json");
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    };

    const importData = (data) => {
        if (confirm('現在のデータをインポートしたデータで上書きしますか？この操作は元に戻せません。')) {
            if (data.userName && data.financialData) {
                state.financialData = processData(data.financialData);
                state.userName = data.userName;
                userNameDisplay.textContent = `${state.userName}の収支`;
            } else {
                // 古い形式のデータの場合
                state.financialData = processData(data);
                state.userName = ''; // 名前はリセット
                userNameDisplay.textContent = '';
            }
            rerender();
            showPopupMessage('データを正常にインポートしました。');
        }
    };

    const handleImportedFileContent = (content) => {
        try {
            let importObject;
            try {
                importObject = JSON.parse(content);
            } catch (e) {
                // パースに失敗した場合、古い形式の暗号化データかもしれない
                try {
                    const bytes = CryptoJS.AES.decrypt(content, ENCRYPTION_KEY);
                    const decryptedData = JSON.parse(bytes.toString(CryptoJS.enc.Utf8));
                    if (confirm('古い形式の暗号化データを検出しました。インポートしますか？')) {
                        importData(decryptedData);
                    }
                } catch (decryptError) {
                    showPopupMessage('無効なファイル形式です。');
                    console.error('JSONのインポートに失敗しました:', e, decryptError);
                }
                return;
            }

            if (importObject.isPasswordProtected) {
                const password = prompt('このファイルはパスワードで保護されています。パスワードを入力してください:');
                if (password === null) return; // キャンセルされた

                try {
                    const bytes = CryptoJS.AES.decrypt(importObject.data, password);
                    const decryptedString = bytes.toString(CryptoJS.enc.Utf8);
                    if (!decryptedString) throw new Error("Decryption failed");
                    const decryptedData = JSON.parse(decryptedString);
                    importData(decryptedData);
                } catch (err) {
                    showPopupMessage('パスワードが違うか、ファイルが破損しています。');
                    console.error('復号化に失敗:', err);
                }
            } else {
                // isPasswordProtected: false またはフラグなしの新しい形式
                importData(importObject.data || importObject); // 古い形式も考慮
            }
        } catch (err) {
            showPopupMessage('無効なファイル形式です。');
            console.error('JSONのインポートに失敗しました:', err);
        }
    };

    const importFromJson = () => {
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = 'application/json';
        input.onchange = e => {
            const file = e.target.files[0];
            if (!file) return;
            const reader = new FileReader();
            reader.onload = readerEvent => {
                handleImportedFileContent(readerEvent.target.result);
            };
            reader.readAsText(file);
        };
        input.click();
    };

    // --- Drag & Drop for Import ---
    appContainer.addEventListener('dragover', (e) => {
        e.preventDefault();
        appContainer.classList.add('drag-over');
    });

    appContainer.addEventListener('dragleave', (e) => {
        e.preventDefault();
        appContainer.classList.remove('drag-over');
    });

    appContainer.addEventListener('drop', (e) => {
        e.preventDefault();
        appContainer.classList.remove('drag-over');

        const files = e.dataTransfer.files;
        if (files.length > 0) {
            const file = files[0];
            if (file.type === 'application/json') {
                const reader = new FileReader();
                reader.onload = readerEvent => {
                    handleImportedFileContent(readerEvent.target.result);
                };
                reader.readAsText(file);
            } else {
                showPopupMessage('JSONファイルのみをドロップしてください。');
            }
        }
    });

    // --- Rendering ---
    const highlightSearchTerm = (text, searchTerm) => {
        if (!searchTerm) return text;
        const escapedSearchTerm = searchTerm.replace(/[.*+?^${}()|[\\]/g, '\\$&');
        const regex = new RegExp(`(${escapedSearchTerm})`, 'gi');
        return text.replace(regex, '<span class="search-highlight">$1</span>');
    };

    const render = () => {
        if (!appContainer) return;
        
        calculateAllBalances();
        const formatter = new Intl.NumberFormat('ja-JP', { style: 'currency', currency: 'JPY' });
        appContainer.innerHTML = '';
        const searchTerm = (state.searchTerm || '').toLowerCase();

        const grandTotalIncome = state.totals.income;
        const grandTotalExpenditure = state.totals.expenditure;
        const netTotal = state.totals.balance;
        const initialBalance = state.totals.initial;

        appContainer.innerHTML = `
        <div class="total-summary-card">
            <h2>総合収支</h2>
            <div class="summary-grid">
                <div class="summary-grid-item">
                    <div class="label">初期残高</div>
                    ${state.isEditMode
                        ? `<input type="number" class="initial-balance-editor" value="${initialBalance}">`
                        : `<div class="value">${formatter.format(initialBalance)}</div>`}
                </div>
                <div class="summary-grid-item"><div class="label">総収入</div><div class="value income">${formatter.format(grandTotalIncome)}</div></div>
                <div class="summary-grid-item"><div class="label">総支出</div><div class="value expenditure">${formatter.format(grandTotalExpenditure)}</div></div>
                <div class="summary-grid-item"><div class="label">最終残高</div><div class="value net">${formatter.format(netTotal)}</div></div>
            </div>
        </div>`;

        const filteredData = state.financialData.map(year => {
            const filteredMonths = year.months.filter(month => {
                if (!searchTerm) return true;
                const inMonth = month.month.toLowerCase().includes(searchTerm);
                const inNote = month.note.toLowerCase().includes(searchTerm);
                const inIncome = month.income.some(i => i.name.toLowerCase().includes(searchTerm));
                const inExpenditure = month.expenditure.some(e => e.name.toLowerCase().includes(searchTerm));
                return inMonth || inNote || inIncome || inExpenditure;
            });
            return {...year, months: filteredMonths};
        }).filter(year => year.months.length > 0);

        if (filteredData.length === 0) {
            appContainer.innerHTML += '<div class="empty-list-placeholder">データがありません。</div>';
            return;
        }

        filteredData.forEach(year => {
            const yearSection = document.createElement('div');
            yearSection.className = 'year-section';

            const yearHeaderContent = state.isEditMode
                ? `<div class="year-header-editor">
                       <input type="text" class="year-title-editor" value="${year.year}" data-year-id="${year.id}">
                       <button class="control-btn edit-months-btn" data-year-id="${year.id}" title="月の編集">${icons.edit}</button>
                       <button class="control-btn add-year-btn" data-after-year-id="${year.id}" title="新しい年度/年を追加">${icons.add}</button>
                   </div>`
                : `<h2 class="year-title">${highlightSearchTerm(year.year, searchTerm)}</h2>`;

            yearSection.innerHTML = `<div class="year-header">${yearHeaderContent}</div><div class="months-grid"></div>`;
            const monthsGrid = yearSection.querySelector('.months-grid');

            year.months.forEach(month => {
                const totalIncome = month.totalIncome;
                const totalExpenditure = month.totalExpenditure;
                const monthlyNet = totalIncome - totalExpenditure;
                const incomePercent = (totalIncome + totalExpenditure) > 0 ? (totalIncome / (totalIncome + totalExpenditure)) * 100 : 0;

                const incomeItems = month.income.length > 0 ? month.income.map(item => `
                    <li data-year-id="${year.id}" data-month-id="${month.id}" data-item-id="${item.id}" data-type="income">
                        <span>${highlightSearchTerm(item.name, searchTerm)}</span><span>${formatter.format(item.amount)}</span>
                        <div class="edit-controls"><button class="control-btn edit-btn">${icons.edit}</button><button class="control-btn delete-btn">${icons.delete}</button></div>
                    </li>`).join('') : '<div class="empty-list-placeholder">収入データがありません</div>';

                const expenditureItems = month.expenditure.length > 0 ? month.expenditure.map(item => `
                     <li data-year-id="${year.id}" data-month-id="${month.id}" data-item-id="${item.id}" data-type="expenditure">
                        <span>${highlightSearchTerm(item.name, searchTerm)}</span><span>${formatter.format(item.amount)}</span>
                        <div class="edit-controls"><button class="control-btn edit-btn">${icons.edit}</button><button class="control-btn delete-btn">${icons.delete}</button></div>
                    </li>`).join('') : '<div class="empty-list-placeholder">支出データがありません</div>';

                const card = document.createElement('div');
                card.className = 'month-card';
                if (monthlyNet > 0) card.classList.add('positive-flow');
                if (monthlyNet < 0) card.classList.add('negative-flow');

                card.innerHTML = `
                    <div class="month-header">
                        <h3 class="month-title">${highlightSearchTerm(month.month, searchTerm)}</h3>
                    </div>
                    <div class="summary">
                        <div class="summary-item"><span class="label">前月残額:</span> <span>${formatter.format(month.startingBalance)}</span></div>
                        <div class="summary-item"><span class="label">収入合計:</span> <span class="value income">${formatter.format(totalIncome)}</span></div>
                        <div class="summary-item"><span class="label">支出合計:</span> <span class="value expenditure">${formatter.format(totalExpenditure)}</span></div>
                    </div>
                    <div class="chart-container" data-year-id="${year.id}" data-month-id="${month.id}"><div class="chart"><div class="chart-income" style="width: ${incomePercent}%"></div><div class="chart-expenditure" style="width: ${100 - incomePercent}%"></div></div></div>
                    <div class="details">
                        <div class="notes-section">
                            <h4>メモ</h4>
                            <div class="notes-content" style="display: ${state.isEditMode ? 'none' : 'block'};">${highlightSearchTerm(month.note || '-', searchTerm)}</div>
                            <textarea class="notes-editor" style="display: ${state.isEditMode ? 'block' : 'none'};" data-year-id="${year.id}" data-month-id="${month.id}">${month.note}</textarea>
                        </div>
                        <div class="income-section"><h4>収入 <div class="edit-controls"><button class="control-btn add-btn" data-year-id="${year.id}" data-month-id="${month.id}" data-type="income">${icons.add}</button></div></h4><ul>${incomeItems}</ul></div>
                        <div class="expenditure-section"><h4>支出 <div class="edit-controls"><button class="control-btn add-btn" data-year-id="${year.id}" data-month-id="${month.id}" data-type="expenditure">${icons.add}</button></div></h4><ul>${expenditureItems}</ul></div>
                    </div>
                    <div class="final-balance">最終残額: ${formatter.format(month.finalBalance)}</div>
                `;
                monthsGrid.appendChild(card);
            });
            appContainer.appendChild(yearSection);
        });

        initTooltips();
    };

    const initTooltips = () => {
        tippy('.chart-container', {
            content: (reference) => {
                const { yearId, monthId } = reference.dataset;
                const month = findMonth(yearId, monthId);
                if (!month) return 'データが見つかりません';

                const total = month.totalIncome + month.totalExpenditure;
                if (total === 0) {
                    return '収入・支出がありません';
                }

                const incomePercent = (month.totalIncome / total) * 100;
                const expenditurePercent = (month.totalExpenditure / total) * 100;

                return `収入: ${incomePercent.toFixed(1)}%<br>支出: ${expenditurePercent.toFixed(1)}%`;
            },
            allowHTML: true,
            theme: 'light',
        });
    };

    const rerender = () => {
        saveState();
        render();
    };

    // --- Event Listeners ---
    if (passwordForm) {
        passwordForm.addEventListener('submit', async (e) => {
            e.preventDefault();
            const enteredPassword = passwordInput.value;
            passwordError.textContent = ''; // Clear previous errors
            await loadAndRenderData(enteredPassword);
        });
    }

    if (newDataButton) {
        newDataButton.addEventListener('click', () => {
            if (startYearInput) {
                startYearInput.value = new Date().getFullYear();
                if (document.querySelector('input[name="creation-type"]:checked').value === 'fiscal') {
                    startYearInput.value = '1年次';
                }
            }
            if (initialBalanceInput) initialBalanceInput.value = 0;
            if (newDataModal) showModal(newDataModal);
        });

        const creationTypeRadios = document.querySelectorAll('input[name="creation-type"]');
        creationTypeRadios.forEach(radio => {
            radio.addEventListener('change', (e) => {
                if (e.target.value === 'fiscal') {
                    startYearInput.value = '1年次';
                } else {
                    startYearInput.value = new Date().getFullYear();
                }
            });
        });
    }

    if (newDataModal) {
        newDataCancelButton.addEventListener('click', () => hideModal(newDataModal));
        newDataCloseButton.addEventListener('click', () => hideModal(newDataModal));
        newDataModal.addEventListener('click', (e) => {
            if (e.target === newDataModal) hideModal(newDataModal);
        });
    }

    if (newDataForm) {
        newDataForm.addEventListener('submit', (e) => {
            e.preventDefault();
            const creationType = document.querySelector('input[name="creation-type"]:checked').value;
            const yearName = startYearInput.value.trim();
            const initialBalance = parseInt(initialBalanceInput.value, 10);
            const userName = userNameInput.value.trim();

            if (!userName) {
                showPopupMessage('名前を入力してください。');
                return;
            }
            if (!yearName) {
                showPopupMessage('有効な年/年度名を入力してください。');
                return;
            }
            if (isNaN(initialBalance)) {
                showPopupMessage('有効な初期残高を入力してください。');
                return;
            }

            if (confirm('現在のデータをすべて削除し、新しいデータを作成します。よろしいですか？この操作は元に戻せません。')) {
                const newData = createNewFinancialData({
                    type: creationType,
                    yearName: yearName,
                    initialBalance: initialBalance
                });
                state.financialData = processData(newData);
                state.userName = userName;
                userNameDisplay.textContent = `${state.userName}の収支`;
                hideModal(newDataModal);
                rerender();
            }
        });
    }

    editModeToggle.addEventListener('click', () => {
        state.isEditMode = !state.isEditMode;
        document.body.classList.toggle('edit-mode', state.isEditMode);
        editModeToggle.textContent = state.isEditMode ? '編集モード終了' : '編集モード';
        editModeToggle.classList.toggle('active', state.isEditMode);

        if (state.isEditMode) {
            userNameDisplay.style.display = 'none';
            userNameEdit.style.display = 'block';
            userNameEdit.value = state.userName;
            userNameEdit.focus();
        } else {
            userNameDisplay.style.display = 'block';
            userNameEdit.style.display = 'none';
            // 保存処理はchangeイベントで行う
        }

        rerender(); // Rerender to show/hide notes editor
    });

    userNameEdit.addEventListener('change', (e) => {
        const newName = e.target.value.trim();
        if (newName) {
            state.userName = newName;
            userNameDisplay.textContent = `${state.userName}の収支`;
            saveState();
        } else {
            e.target.value = state.userName; // 空の場合は元に戻す
            showPopupMessage('名前は空にできません。');
        }
    });

    resetDataButton.addEventListener('click', async () => {
        if (confirm('本当にすべてのデータをリセットしますか？この操作は元に戻せません。')) {
            localStorage.removeItem('financialData');
            localStorage.removeItem('userName');
            localStorage.removeItem('password');
            state.userName = '';
            state.password = '';
            userNameDisplay.textContent = '';
            state.financialData = await loadDefaultData();
            showPopupMessage('データが正常にリセットされました。');
            rerender();
        }
    });

    summaryButton.addEventListener('click', renderSummaryModal);
    summaryCloseButton.addEventListener('click', () => hideModal(summaryModal));
    summaryModal.addEventListener('click', (e) => { if(e.target === summaryModal) hideModal(summaryModal); });

    if (categorySummaryButton && categorySummaryModal && categorySummaryCloseButton) {
        categorySummaryButton.addEventListener('click', renderCategorySummaryModal);
        categorySummaryCloseButton.addEventListener('click', () => hideModal(categorySummaryModal));
        categorySummaryModal.addEventListener('click', (e) => { if (e.target === categorySummaryModal) hideModal(categorySummaryModal); });
    }

    if (settingsButton) {
        settingsButton.addEventListener('click', () => {
            categoryInput.value = state.categories.join('\n');
            passwordSettingInput.value = state.password;
            showModal(settingsModal);


        });
    }

    if (settingsCloseButton) {
        settingsCloseButton.addEventListener('click', () => hideModal(settingsModal));
    }

    if (settingsModal) {
        settingsModal.addEventListener('click', (e) => {
            if (e.target === settingsModal) hideModal(settingsModal);
        });
    }

    if (cancelCategoriesButton) {
        cancelCategoriesButton.addEventListener('click', () => hideModal(settingsModal));
    }

    if (saveSettingsButton) {
        saveSettingsButton.addEventListener('click', () => {
            const newCategories = categoryInput.value.split('\n').map(cat => cat.trim()).filter(cat => cat !== '');
            state.categories = Array.from(new Set(newCategories)); // 重複を排除
            
            const newPassword = passwordSettingInput.value;
            if (newPassword !== state.password) {
                if (newPassword === "") {
                    if (confirm("パスワードを削除すると、データの暗号化が解除されます。よろしいですか？")) {
                        state.password = "";
                        showPopupMessage('パスワードが削除されました。');
                    }
                } else {
                    state.password = newPassword;
                    showPopupMessage('パスワードが設定/更新されました。');
                }
            }

            hideModal(settingsModal);
            rerender();
            showPopupMessage('設定を保存しました。');
        });
    }

    if (autoGetCategoriesButton) {
        autoGetCategoriesButton.addEventListener('click', () => {
            // Get existing categories from the textarea
            const existingCategories = categoryInput.value.split('\n').map(cat => cat.trim()).filter(cat => cat);

            // Get categories from financial data
            const extractedCategories = extractCategoriesFromFinancialData();

            // Combine them and get unique values
            const combined = new Set([...existingCategories, ...extractedCategories]);

            // Sort and update the textarea
            categoryInput.value = Array.from(combined).sort().join('\n');
        });
    }

    exportJsonButton.addEventListener('click', exportToJson);
    importJsonButton.addEventListener('click', importFromJson);

    searchBox.addEventListener('input', (e) => {
        state.searchTerm = e.target.value;
        render();
    });

    appContainer.addEventListener('click', e => {
        if (!state.isEditMode) return;
        const addBtn = e.target.closest('.add-btn');
        const editBtn = e.target.closest('.edit-btn');
        const deleteBtn = e.target.closest('.delete-btn');
        const addYearBtn = e.target.closest('.add-year-btn');

        if (addBtn) handleAddItem(addBtn.dataset.yearId, addBtn.dataset.monthId, addBtn.dataset.type);
        else if (editBtn) {
            const li = e.target.closest('li');
            handleEditItem(li.dataset.yearId, li.dataset.monthId, li.dataset.type, li.dataset.itemId);
        } else if (deleteBtn) {
            const li = e.target.closest('li');
            handleDeleteItem(li.dataset.yearId, li.dataset.monthId, li.dataset.type, li.dataset.itemId);
        } else if (addYearBtn) {
            const afterYearId = addYearBtn.dataset.afterYearId;
            openAddYearModal(afterYearId);
        } else if (e.target.closest('.edit-months-btn')) {
            const editMonthsBtn = e.target.closest('.edit-months-btn');
            const yearId = editMonthsBtn.dataset.yearId;
            openEditMonthsModal(yearId);
        }
    });

    appContainer.addEventListener('change', e => {
        if (!state.isEditMode) return;

        if (e.target.classList.contains('notes-editor')) {
            const { yearId, monthId } = e.target.dataset;
            handleNoteUpdate(yearId, monthId, e.target.value);
        } else if (e.target.classList.contains('year-title-editor')) {
            const { yearId } = e.target.dataset;
            const newYearName = e.target.value.trim();
            const year = findYear(yearId);

            if (newYearName && year) {
                handleYearNameUpdate(yearId, newYearName);
            } else if (year) {
                e.target.value = year.year; // Restore original value
                showPopupMessage('年度名は空にできません。');
            }
        } else if (e.target.classList.contains('initial-balance-editor')) {
            const newInitialBalance = parseInt(e.target.value, 10);
            if (!isNaN(newInitialBalance)) {
                // 最初の月のstartingBalanceを更新
                if (state.financialData.length > 0 && state.financialData[0].months.length > 0) {
                    state.financialData[0].months[0].startingBalance = newInitialBalance;
                    rerender();
                }
            } else {
                showPopupMessage('有効な初期残高を入力してください。');
                rerender(); // 無効な入力の場合は元の値を再表示
            }
        }
    });

    modalForm.addEventListener('submit', e => {
        e.preventDefault();
        const { yearId, monthId, type, itemId } = state.editingItem;
        const name = itemNameInput.value.trim();
        const amount = parseInt(itemAmountInput.value, 10);
        if (!name || isNaN(amount)) return showPopupMessage('有効な項目名と金額を入力してください。');
        
        // Add new item name to categories if it doesn't exist
        if (!state.categories.includes(name)) {
            state.categories.push(name);
            state.categories.sort(); // ソートして表示順を整える
        }

        const month = findMonth(yearId, monthId);
        if (itemId) {
            const item = month[type].find(i => i.id === itemId);
            item.name = name; item.amount = amount;
        } else {
            month[type].push({ id: `item_${Date.now()}`, name, amount });
        }
        hideModal(modal);
        rerender();
    });

    cancelButton.addEventListener('click', () => hideModal(modal));

    // --- Add Year Modal ---
    let currentAfterYearId = null;

    const populateMonthCheckboxes = (isFiscal) => {
        const months = isFiscal
            ? ['4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月', '1月', '2月', '3月']
            : ['1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月'];
        
        addYearMonthsContainer.innerHTML = months.map(month => `
            <div class="month-checkbox-item">
                <input type="checkbox" id="month-check-${month}" value="${month}" checked>
                <label for="month-check-${month}">${month}</label>
            </div>
        `).join('');
    };

    const openAddYearModal = (afterYearId) => {
        currentAfterYearId = afterYearId;
        const afterYear = findYear(afterYearId);
        const currentYearName = afterYear ? afterYear.year : '';
        
        let nextYearName = '';
        const nendoMatch = currentYearName.match(/(\d+)年次/);
        const yearMatch = currentYearName.match(/(\d{4})/);

        let isFiscalType = false;
        if (afterYear) {
            isFiscalType = afterYear.year.includes('年度') || afterYear.year.includes('年次');
        }

        if (nendoMatch) {
            const yearNumber = parseInt(nendoMatch[1], 10);
            nextYearName = `${yearNumber + 1}年次`;
        } else if (yearMatch) {
            const yearNumber = parseInt(yearMatch[1], 10);
            nextYearName = currentYearName.replace(String(yearNumber), String(yearNumber + 1));
        } else {
            nextYearName = isFiscalType ? `${new Date().getFullYear()}年度` : new Date().getFullYear().toString();
        }

        populateMonthCheckboxes(isFiscalType);
        newYearNameInput.value = nextYearName;
        showModal(addYearModal);
    };

    if (addYearForm) {
        addYearForm.addEventListener('submit', (e) => {
            e.preventDefault();
            const name = newYearNameInput.value.trim();
            if (!name) {
                showPopupMessage('年度/年 名前を入力してください。');
                return;
            }
            
            const afterYear = findYear(currentAfterYearId);
            const isFiscalType = afterYear ? (afterYear.year.includes('年度') || afterYear.year.includes('年次')) : false;
            const type = isFiscalType ? 'fiscal' : 'calendar';

            const selectedMonths = Array.from(addYearMonthsContainer.querySelectorAll('input[type="checkbox"]:checked')).map(cb => cb.value);

            if (selectedMonths.length === 0) {
                showPopupMessage('少なくとも1つの月を選択してください。');
                return;
            }

            handleAddYear({ name, type, selectedMonths, afterYearId: currentAfterYearId });
            hideModal(addYearModal);
        });
    }

    if (addYearModal) {
        addYearCloseButton.addEventListener('click', () => hideModal(addYearModal));
        addYearCancelButton.addEventListener('click', () => hideModal(addYearModal));
        addYearModal.addEventListener('click', (e) => {
            if (e.target === addYearModal) hideModal(addYearModal);
        });
    }

    // --- Edit Months Modal ---
    let currentEditingYearId = null;

    const openEditMonthsModal = (yearId) => {
        currentEditingYearId = yearId;
        const year = findYear(yearId);
        if (!year) return;

        editMonthsTitle.textContent = `「${year.year}」の月の編集`;

        const isFiscal = year.year.includes('年度') || year.year.includes('年次');
        const allMonths = isFiscal
            ? ['4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月', '1月', '2月', '3月']
            : ['1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月'];
        
        const existingMonthNames = new Set(year.months.map(m => m.month));

        editMonthsContainer.innerHTML = allMonths.map(monthName => {
            const monthData = year.months.find(m => m.month === monthName);
            const hasData = monthData && (monthData.income.length > 0 || monthData.expenditure.length > 0 || monthData.note);
            const isChecked = existingMonthNames.has(monthName);

            return `
                <div class="month-checkbox-item" title="${hasData ? 'データが含まれているため削除できません' : ''}">
                    <input type="checkbox" id="edit-month-${monthName}" value="${monthName}" ${isChecked ? 'checked' : ''} ${hasData ? 'disabled' : ''}>
                    <label for="edit-month-${monthName}" class="${hasData ? 'disabled' : ''}">${monthName}</label>
                </div>
            `;
        }).join('');

        showModal(editMonthsModal);
    };

    if (editMonthsForm) {
        editMonthsForm.addEventListener('submit', (e) => {
            e.preventDefault();
            const year = findYear(currentEditingYearId);
            if (!year) return;

            const selectedMonthNames = new Set(Array.from(editMonthsContainer.querySelectorAll('input[type="checkbox"]:checked')).map(cb => cb.value));
            const existingMonthNames = new Set(year.months.map(m => m.month));

            // Remove months
            year.months = year.months.filter(month => selectedMonthNames.has(month.month));

            // Add new months
            selectedMonthNames.forEach(monthName => {
                if (!existingMonthNames.has(monthName)) {
                    year.months.push({
                        id: `month_${Date.now()}_${Math.random()}`,
                        month: monthName,
                        note: '',
                        startingBalance: 0, // Will be recalculated
                        income: [],
                        expenditure: []
                    });
                }
            });

            hideModal(editMonthsModal);
            rerender();
        });
    }

    if (deleteYearButton) {
        deleteYearButton.addEventListener('click', () => {
            if (!currentEditingYearId) return;
            if (confirm('この年を完全に削除しますか？この操作は元に戻せません。')) {
                state.financialData = state.financialData.filter(year => year.id !== currentEditingYearId);
                hideModal(editMonthsModal);
                rerender();
            }
        });
    }

    if (editMonthsModal) {
        editMonthsCloseButton.addEventListener('click', () => hideModal(editMonthsModal));
        editMonthsCancelButton.addEventListener('click', () => hideModal(editMonthsModal));
        editMonthsModal.addEventListener('click', (e) => {
            if (e.target === editMonthsModal) hideModal(editMonthsModal);
        });
    }

    // --- Bulk Edit ---
    const bulkEditMonthsContainer = document.getElementById('bulk-edit-months');

    const populateBulkEditMonths = () => {
        let content = '';
        const createCheckbox = (id, label, value, isChecked = true, customClass = '') => `
            <div class="month-checkbox-item ${customClass}">
                <input type="checkbox" id="${id}" value="${value}" ${isChecked ? 'checked' : ''}>
                <label for="${id}">${label}</label>
            </div>`;

        content += `<div class="month-checkbox-group">
            ${createCheckbox('bulk-month-all', 'すべて選択/解除', 'all', true, 'select-all-global')}
        </div>`;

        state.financialData.forEach(year => {
            content += `<div class="month-checkbox-group" data-year-id="${year.id}">`;
            content += `<span class="month-checkbox-group-title">${year.year}</span>`;
            const yearCheckboxId = `bulk-month-year-${year.id}`;
            content += createCheckbox(yearCheckboxId, 'この年の月をすべて選択', year.id, true, 'select-all-year');
            
            year.months.forEach(month => {
                content += createCheckbox(`bulk-month-${month.id}`, month.month, month.id, true);
            });
            content += `</div>`;
        });
        bulkEditMonthsContainer.innerHTML = content;
    };

    bulkEditButton.addEventListener('click', () => {
        populateBulkEditSelect();
        populateBulkEditMonths();
        showModal(bulkEditModal);
    });

    bulkCancelButton.addEventListener('click', () => hideModal(bulkEditModal));
    bulkEditModal.addEventListener('click', (e) => { 
        if(e.target === bulkEditModal) hideModal(bulkEditModal); 
    });

    bulkEditMonthsContainer.addEventListener('change', (e) => {
        const target = e.target;
        if (target.type !== 'checkbox') return;
        const isChecked = target.checked;

        if (target.value === 'all') {
            bulkEditMonthsContainer.querySelectorAll('input[type="checkbox"]').forEach(cb => cb.checked = isChecked);
        } else if (target.parentElement.classList.contains('select-all-year')) {
            const yearGroup = target.closest('.month-checkbox-group');
            yearGroup.querySelectorAll('input[type="checkbox"]').forEach(cb => cb.checked = isChecked);
            if (!isChecked) {
                document.getElementById('bulk-month-all').checked = false;
            }
        } else {
            if (!isChecked) {
                document.getElementById('bulk-month-all').checked = false;
                const yearCheckbox = target.closest('.month-checkbox-group').querySelector('.select-all-year input');
                if (yearCheckbox) yearCheckbox.checked = false;
            }
        }
    });

    bulkEditForm.addEventListener('submit', e => {
        e.preventDefault();
        const oldName = bulkEditSelect.value;
        const newName = bulkEditNewName.value.trim();
        const newAmountStr = bulkEditNewAmount.value;
        const newAmount = newAmountStr ? parseInt(newAmountStr, 10) : null;

        const selectedMonthIds = Array.from(bulkEditMonthsContainer.querySelectorAll('input[type="checkbox"]:checked'))
            .map(cb => cb.value)
            .filter(value => value !== 'all' && !value.startsWith('year_'));

        if (!oldName) return showPopupMessage('編集する項目を選択してください。');
        if (newName === '' && newAmount === null) return showPopupMessage('新しい項目名または新しい金額の少なくとも一方を入力してください。');
        if (newAmount !== null && isNaN(newAmount)) return showPopupMessage('有効な金額を入力してください。');
        if (selectedMonthIds.length === 0) return showPopupMessage('適用する月を少なくとも1つ選択してください。');

        let updatedCount = 0;
        state.financialData.forEach(year => {
            year.months.forEach(month => {
                if (selectedMonthIds.includes(month.id)) {
                    const updateItems = (items) => {
                        items.forEach(item => {
                            if (item.name === oldName) {
                                if (newName !== '') item.name = newName;
                                if (newAmount !== null) item.amount = newAmount;
                                updatedCount++;
                            }
                        });
                    };
                    updateItems(month.income);
                    updateItems(month.expenditure);
                }
            });
        });

        if (updatedCount > 0) {
            if (confirm(`${updatedCount}個の項目を更新します。よろしいですか？`)) {
                bulkEditNewName.value = '';
                bulkEditNewAmount.value = '';
                hideModal(bulkEditModal);
                rerender();
            }
        } else {
            showPopupMessage('選択された月に該当する項目が見つかりませんでした。');
            hideModal(bulkEditModal);
        }
    });

    const populateBulkEditSelect = () => {
        const allItems = new Set();
        state.financialData.forEach(year => {
            year.months.forEach(month => {
                month.income.forEach(item => allItems.add(item.name));
                month.expenditure.forEach(item => allItems.add(item.name));
            });
        });

        bulkEditSelect.innerHTML = '<option value="">項目を選択してください</option>';
        const sortedItems = Array.from(allItems).sort();
        sortedItems.forEach(name => {
            const option = document.createElement('option');
            option.value = name;
            option.textContent = name;
            bulkEditSelect.appendChild(option);
        });
    };

    // --- Initial Load ---
    initializeApp();
});

// ポップアップメッセージを表示する関数
function showPopupMessage(message) {
    const popupMessage = document.getElementById('popupMessage');
    const popupText = document.getElementById('popupText');
    const closePopup = document.getElementById('closePopup');

    popupText.textContent = message;
    popupMessage.style.display = 'flex'; // ポップアップを表示

    closePopup.onclick = () => {
        popupMessage.style.display = 'none'; // ポップアップを非表示
    };

    // ポップアップの外側をクリックしても閉じるようにする
    popupMessage.onclick = (event) => {
        if (event.target === popupMessage) {
            popupMessage.style.display = 'none';
        }
    };
}
