class SeatingApp {
    constructor() {
        this.students = [];
        this.seats = [];
        this.rows = 6;
        this.cols = 8;
        this.selectedSeats = new Set(); // 多选座位集合
        this.isSelecting = false; // 是否正在框选
        this.selectionStart = null; // 框选起始点
        this.selectionBox = null; // 选择框元素
        this.history = [];
        this.historyIndex = -1;
        this.constraints = [];
        this.showCoordinates = false;
        this.selectedFont = 'song';
        this.maleColor = '#2563eb';
        this.femaleColor = '#ec4899';
        this.isDragging = false; // 单选拖拽状态标志
        this.isMultiDragging = false; // 多选拖拽状态标志
        this.searchDebounceTimer = null; // 搜索防抖计时器
        
        // 触摸拖拽状态
        this.touchDragData = null; // 触摸拖拽的数据
        this.touchDragElement = null; // 触摸拖拽的可视化元素
        this.isTouchDragging = false; // 是否正在触摸拖拽
        this.touchStartTime = 0; // 触摸开始时间（用于区分点击和拖拽）
        
        this.init();
    }

    init() {
        this.loadData();
        this.setupEventListeners();
        this.setupOrientationHandling(); // 横竖屏切换处理
        this.renderClassroom();
        this.renderStudentList();
        this.updateStats();
        this.applyCurrentFilter();
        this.updateHistoryButtons();
        this.initializeLayoutSettings();
    }

    loadData() {
        const savedData = localStorage.getItem('seatingData');
        if (savedData) {
            const data = JSON.parse(savedData);
            this.students = data.students || [];
            this.rows = data.rows || 6;
            this.cols = data.cols || 8;
            this.constraints = data.constraints || [];
            this.showCoordinates = data.showCoordinates !== undefined ? data.showCoordinates : false;
            this.selectedFont = data.selectedFont || 'song';
            this.maleColor = data.maleColor || '#2563eb';
            this.femaleColor = data.femaleColor || '#ec4899';
            // 恢复历史记录
            this.history = data.history || [];
            this.historyIndex = data.historyIndex !== undefined ? data.historyIndex : -1;
            
            // 直接恢复保存的座位数据
            if (data.seats) {
                this.seats = data.seats;
                return; // 如果已经恢复了座位数据，不需要再初始化座位
            }
        }
        this.initializeSeats();
        
        // 在加载完数据后，如果没有历史记录，添加初始状态
        if (this.history.length === 0) {
            this.addToHistory('seatArrangement', { seats: this.seats });
            this.historyIndex = -1; // 重置为初始状态，因为这是起始点
        }
    }

    saveData() {
        const data = {
            students: this.students,
            seats: this.seats,
            rows: this.rows,
            cols: this.cols,
            constraints: this.constraints,
            showCoordinates: this.showCoordinates,
            selectedFont: this.selectedFont,
            maleColor: this.maleColor,
            femaleColor: this.femaleColor,
            history: this.history,
            historyIndex: this.historyIndex
        };
        localStorage.setItem('seatingData', JSON.stringify(data));
    }

    initializeSeats() {
        // 保存现有座位的删除状态和学生信息
        const existingSeatData = {};
        if (this.seats) {
            this.seats.forEach(seat => {
                existingSeatData[seat.id] = {
                    student: seat.student,
                    isDeleted: seat.isDeleted || false
                };
            });
        }
        
        this.seats = [];
        // 教室坐标系统规则（以讲台为中心）:
        // - 讲台位于教室前方，面向学生
        // - 第一排：最靠近讲台的座位行
        // - 第一列：从讲台看向学生的最左侧列
        // 
        // 内部坐标系统:
        // - row: 0=第一排(靠近讲台), 数值增加表示远离讲台
        // - col: 0=第一列(最左边), 数值增加表示向右移动
        // 
        // 显示坐标系统:
        // - 第一排第一列显示为 "1-1"
        // - 按学号排座时：从第一排第一列开始，按行优先从左到右排列
        for (let row = 0; row < this.rows; row++) {
            for (let col = 0; col < this.cols; col++) {
                const seatId = `${row}-${col}`;
                const existingData = existingSeatData[seatId];
                
                this.seats.push({
                    id: seatId,
                    row: row,    // 内部行索引: 0-based, 0=第一排(靠近讲台)
                    col: col,    // 内部列索引: 0-based, 0=第一列(最左边)
                    student: existingData ? existingData.student : null,
                    position: row * this.cols + col + 1,  // 线性位置编号
                    isDeleted: existingData ? existingData.isDeleted : false  // 座位删除状态
                });
            }
        }
    }

    setupEventListeners() {
        document.getElementById('addStudent').addEventListener('click', () => {
            this.showStudentModal();
        });

        document.getElementById('randomSeat').addEventListener('click', () => {
            this.randomSeatArrangement();
        });

        document.getElementById('clearSeats').addEventListener('click', () => {
            this.clearAllSeats();
        });

        document.getElementById('saveLayout').addEventListener('click', () => {
            this.saveCurrentLayout();
        });


        document.getElementById('printLayout').addEventListener('click', () => {
            this.printLayout();
        });

        // 撤销按钮在主页工具栏中，直接绑定
        document.getElementById('undoBtn').addEventListener('click', () => {
            this.undo();
        });

        // 使用事件委托处理模态框和下拉框中的按钮
        document.addEventListener('click', (e) => {
            switch (e.target.id) {
                case 'applyLayout':
                    this.applyNewLayout();
                    break;
                case 'applyLayoutDropdown':
                    this.applyNewLayoutFromDropdown();
                    break;
                case 'addConstraint':
                    this.addConstraint();
                    break;
            }
        });

        document.getElementById('saveStudent').addEventListener('click', () => {
            this.saveStudent();
        });

        document.getElementById('cancelStudent').addEventListener('click', () => {
            this.hideStudentModal();
        });

        document.querySelector('.modal-close').addEventListener('click', () => {
            this.hideStudentModal();
        });

        document.getElementById('studentModal').addEventListener('click', (e) => {
            if (e.target.id === 'studentModal') {
                this.hideStudentModal();
            }
        });

        document.getElementById('searchStudent').addEventListener('input', (e) => {
            // 使用防抖优化搜索性能
            if (this.searchDebounceTimer) {
                clearTimeout(this.searchDebounceTimer);
            }
            const searchValue = e.target.value;
            this.searchDebounceTimer = setTimeout(() => {
                this.filterStudents(searchValue);
            }, 300); // 300ms 延迟
        });

        document.getElementById('filterStudents').addEventListener('change', (e) => {
            this.filterStudentsByStatus(e.target.value);
        });

        document.getElementById('clearAllStudents').addEventListener('click', () => {
            this.clearAllStudents();
        });

        document.getElementById('seatingSettingsBtn').addEventListener('click', () => {
            this.showSeatingSettingsModal();
        });



        document.getElementById('closeSeatingSettings').addEventListener('click', () => {
            this.hideSeatingSettingsModal();
        });

        document.getElementById('applySeatingRules').addEventListener('click', () => {
            this.ruleBasedSeatArrangement();
        });

        document.getElementById('layoutSettingsBtn').addEventListener('click', (e) => {
            e.stopPropagation();
            this.toggleLayoutSettingsDropdown();
        });

        document.getElementById('showCoordinatesToggle').addEventListener('change', (e) => {
            this.toggleCoordinatesDisplay(e.target.checked);
        });

        // Close dropdown when clicking outside
        document.addEventListener('click', (e) => {
            const settingsContainer = e.target.closest('.layout-settings-container');
            if (!settingsContainer) {
                this.hideLayoutSettingsDropdown();
            }
        });

        document.getElementById('fontSelectDropdown').addEventListener('change', (e) => {
            this.changeFontFamily(e.target.value);
        });

        // Color picker event listeners
        document.addEventListener('click', (e) => {
            if (e.target.classList.contains('color-swatch')) {
                this.selectColor(e.target);
            }
        });

        // 座位轮换事件监听器
        document.getElementById('rotateRowLeft').addEventListener('click', () => {
            this.rotateSeats('rowLeft');
        });

        document.getElementById('rotateRowRight').addEventListener('click', () => {
            this.rotateSeats('rowRight');
        });

        document.getElementById('rotateColForward').addEventListener('click', () => {
            this.rotateSeats('colForward');
        });

        document.getElementById('rotateColBackward').addEventListener('click', () => {
            this.rotateSeats('colBackward');
        });

        // 添加约束输入框的回车键支持
        document.getElementById('constraintInput').addEventListener('keypress', (e) => {
            if (e.key === 'Enter') {
                this.addConstraint();
            }
        });

        document.getElementById('seatingSettingsModal').addEventListener('click', (e) => {
            if (e.target.id === 'seatingSettingsModal') {
                this.hideSeatingSettingsModal();
            }
        });

        this.setupExcelEventListeners();
        this.setupMultiSelectEventListeners();
    }

    setupExcelEventListeners() {
        document.getElementById('importExcel').addEventListener('click', () => {
            this.importExcelFile();
        });

        document.getElementById('excelFileInput').addEventListener('change', (e) => {
            this.handleExcelFileSelect(e);
        });

        document.getElementById('cancelImport').addEventListener('click', () => {
            this.hideExcelPreviewModal();
        });

        document.getElementById('confirmImport').addEventListener('click', () => {
            this.confirmExcelImport();
        });

        document.getElementById('excelPreviewModal').addEventListener('click', (e) => {
            if (e.target.id === 'excelPreviewModal') {
                this.hideExcelPreviewModal();
            }
        });

        document.querySelectorAll('.tab-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                this.switchTab(e.target.dataset.tab);
            });
        });
    }

    addToHistory(action, data) {
        // 截取当前索引之后的历史记录（撤销后的重做分支会被丢弃）
        this.history = this.history.slice(0, this.historyIndex + 1);
        
        // 添加新的历史记录
        this.history.push({
            action: action,
            data: JSON.parse(JSON.stringify(data)),
            timestamp: Date.now()
        });
        
        // 限制历史记录最多保存2次（用户可撤销2次）
        const MAX_HISTORY_SIZE = 2;
        if (this.history.length > MAX_HISTORY_SIZE) {
            // 删除最旧的记录，保留最新的2条
            const removeCount = this.history.length - MAX_HISTORY_SIZE;
            this.history = this.history.slice(removeCount);
            // 调整索引位置
            this.historyIndex = this.history.length - 1;
        } else {
            this.historyIndex++;
        }
        
        // 更新撤销/重做按钮状态
        this.updateHistoryButtons();
        
        // 保存到LocalStorage
        this.saveData();
    }

    undo() {
        if (this.historyIndex >= 0) {
            const historyItem = this.history[this.historyIndex];
            this.restoreFromHistory(historyItem);
            this.historyIndex--;
        }
    }

    redo() {
        if (this.historyIndex < this.history.length - 1) {
            this.historyIndex++;
            const historyItem = this.history[this.historyIndex];
            this.restoreFromHistory(historyItem);
        }
    }

    restoreFromHistory(historyItem) {
        switch (historyItem.action) {
            case 'seatArrangement':
                this.seats = JSON.parse(JSON.stringify(historyItem.data.seats));
                this.renderClassroom();
                this.renderStudentList();
                this.updateStats();
                break;
        }
        this.updateHistoryButtons();
    }

    updateHistoryButtons() {
        const undoBtn = document.getElementById('undoBtn');
        
        const canUndo = this.historyIndex >= 0;
        
        if (undoBtn) {
            undoBtn.disabled = !canUndo;
                    }
    }

    setupOrientationHandling() {
        // 处理横竖屏切换
        let orientationChangeTimer = null;
        
        const handleOrientationChange = () => {
            // 清除之前的定时器
            if (orientationChangeTimer) {
                clearTimeout(orientationChangeTimer);
            }
            
            // 延迟执行，等待屏幕旋转动画完成
            orientationChangeTimer = setTimeout(() => {
                // 重新渲染座位图以适应新的屏幕方向
                this.renderClassroom();
                
                // 清除任何选中状态，避免布局错乱
                this.clearSelection();
                
                // 更新滚动位置（移动端可能需要）
                const classroomGrid = document.getElementById('classroomGrid');
                if (classroomGrid && classroomGrid.parentElement) {
                    classroomGrid.parentElement.scrollTop = 0;
                }
                
                console.log('屏幕方向已改变，布局已更新');
            }, 300); // 延迟300ms等待旋转动画完成
        };
        
        // 监听orientationchange事件（旧版浏览器）
        window.addEventListener('orientationchange', handleOrientationChange);
        
        // 监听resize事件作为备选（更通用）
        let resizeTimer = null;
        window.addEventListener('resize', () => {
            if (resizeTimer) {
                clearTimeout(resizeTimer);
            }
            resizeTimer = setTimeout(() => {
                // 检查是否真的是横竖屏切换（而不是键盘弹出等）
                const isLandscape = window.innerWidth > window.innerHeight;
                const wasLandscape = this.lastOrientation === 'landscape';
                
                if (isLandscape !== wasLandscape) {
                    this.lastOrientation = isLandscape ? 'landscape' : 'portrait';
                    handleOrientationChange();
                }
            }, 200);
        });
        
        // 初始化方向状态
        this.lastOrientation = window.innerWidth > window.innerHeight ? 'landscape' : 'portrait';
        
        // Screen Orientation API（现代浏览器）
        if (screen.orientation) {
            screen.orientation.addEventListener('change', () => {
                handleOrientationChange();
            });
        }
    }

    showStudentModal(student = null) {
        const modal = document.getElementById('studentModal');
        const form = document.getElementById('studentForm');
        const title = document.getElementById('modalTitle');
        
        if (student) {
            title.textContent = '编辑学生';
            document.getElementById('studentName').value = student.name;
            document.getElementById('studentId').value = student.id || '';
            document.getElementById('studentHeight').value = student.height || '';
            document.getElementById('studentGender').value = student.gender || '';
            document.getElementById('studentVision').checked = student.needsFrontSeat || false;
            document.getElementById('studentNotes').value = student.notes || '';
            form.dataset.editId = student.uuid;
        } else {
            title.textContent = '添加学生';
            form.reset();
            delete form.dataset.editId;
        }
        
        modal.style.display = 'flex';
    }

    hideStudentModal() {
        document.getElementById('studentModal').style.display = 'none';
    }

    saveStudent() {
        const form = document.getElementById('studentForm');
        const name = document.getElementById('studentName').value.trim();
        
        if (!name) {
            alert('请输入学生姓名');
            return;
        }

        const student = {
            uuid: form.dataset.editId || this.generateUUID(),
            name: name,
            id: document.getElementById('studentId').value.trim(),
            height: parseInt(document.getElementById('studentHeight').value) || null,
            gender: document.getElementById('studentGender').value,
            needsFrontSeat: document.getElementById('studentVision').checked,
            notes: document.getElementById('studentNotes').value.trim(),
            seatId: null
        };

        if (form.dataset.editId) {
            const index = this.students.findIndex(s => s.uuid === form.dataset.editId);
            if (index !== -1) {
                this.students[index] = student;
            }
        } else {
            this.students.push(student);
        }

        this.saveData();
        this.renderStudentList();
        this.updateStats();
        this.applyCurrentFilter();
        this.hideStudentModal();
    }

    deleteStudent(uuid) {
        if (confirm('确定要删除这个学生吗？')) {
            const seat = this.seats.find(s => s.student && s.student.uuid === uuid);
            if (seat) {
                seat.student = null;
            }
            
            this.students = this.students.filter(s => s.uuid !== uuid);
            this.saveData();
            this.renderStudentList();
            this.renderClassroom();
            this.updateStats();
            this.applyCurrentFilter();
        }
    }

    generateUUID() {
        return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
            const r = Math.random() * 16 | 0;
            const v = c == 'x' ? r : (r & 0x3 | 0x8);
            return v.toString(16);
        });
    }

    parseDisplayCoordinate(coordString) {
        if (!coordString || typeof coordString !== 'string') {
            return null;
        }

        const trimmedCoord = coordString.trim();
        const match = trimmedCoord.match(/^(\d+)-(\d+)$/);
        
        if (!match) {
            console.warn('无效的座位坐标格式:', coordString);
            return null;
        }

        const displayRow = parseInt(match[1]);
        const displayCol = parseInt(match[2]);

        const internalRow = this.rows - displayRow;
        const internalCol = displayCol - 1;

        if (internalRow < 0 || internalRow >= this.rows || internalCol < 0 || internalCol >= this.cols) {
            console.warn('座位坐标超出范围:', coordString, '有效范围: 1-' + this.rows + ', 1-' + this.cols);
            return null;
        }

        return {
            row: internalRow,
            col: internalCol,
            displayRow: displayRow,
            displayCol: displayCol,
            seatId: `${internalRow}-${internalCol}`
        };
    }

    renderStudentList() {
        const container = document.getElementById('studentList');
        
        // 优化：使用增量更新而不是完全重建DOM
        // 创建学生UUID映射以便快速查找
        const existingItems = new Map();
        Array.from(container.children).forEach(item => {
            const uuid = item.dataset.studentUuid;
            if (uuid) {
                existingItems.set(uuid, item);
            }
        });
        
        // 创建DocumentFragment减少重排
        const fragment = document.createDocumentFragment();
        const currentStudentUuids = new Set();
        
        this.students.forEach(student => {
            currentStudentUuids.add(student.uuid);
            
            // 检查学生项是否已存在
            let item = existingItems.get(student.uuid);
            const isSeated = this.seats.some(seat => seat.student && seat.student.uuid === student.uuid);
            
            // 如果元素已存在，只更新必要的部分
            if (item) {
                // 更新座位状态类
                if (isSeated && !item.classList.contains('seated')) {
                    item.classList.add('seated');
                } else if (!isSeated && item.classList.contains('seated')) {
                    item.classList.remove('seated');
                }
                
                // 检查内容是否需要更新（比较关键数据）
                const needsUpdate = item.dataset.studentData !== JSON.stringify({
                    name: student.name,
                    gender: student.gender,
                    needsFrontSeat: student.needsFrontSeat,
                    notes: student.notes
                });
                
                if (needsUpdate) {
                    this.updateStudentItemContent(item, student, isSeated);
                }
                
                existingItems.delete(student.uuid);
            } else {
                // 创建新的学生项
                item = this.createStudentItem(student, isSeated);
                fragment.appendChild(item);
            }
        });
        
        // 移除不再存在的学生项
        existingItems.forEach(item => {
            container.removeChild(item);
        });
        
        // 添加新创建的项
        if (fragment.children.length > 0) {
            container.appendChild(fragment);
        }
    }
    
    createStudentItem(student, isSeated) {
        const item = document.createElement('div');
        item.className = 'student-item';
        item.draggable = true;
        item.dataset.studentUuid = student.uuid;
        
        if (isSeated) {
            item.classList.add('seated');
        }
        
        this.updateStudentItemContent(item, student, isSeated);
        
        // 添加拖拽事件监听器
        item.addEventListener('dragstart', (e) => {
            e.dataTransfer.setData('text/plain', student.uuid);
            item.classList.add('dragging');
        });
        
        item.addEventListener('dragend', () => {
            item.classList.remove('dragging');
        });
        
        return item;
    }
    
    updateStudentItemContent(item, student, isSeated) {
        // 构建学生详情信息
        let details = [];
        if (student.gender) {
            const genderText = student.gender === 'male' ? '男' : student.gender === 'female' ? '女' : student.gender;
            details.push(`性别: ${genderText}`);
        }
        if (student.needsFrontSeat) details.push('需前排');
        
        item.innerHTML = `
            <div class="student-info">
                <div class="student-name">${student.name}</div>
                <div class="student-details">
                    ${details.join(' | ')}
                </div>
                ${student.notes ? `<div class="student-notes">备注: ${student.notes}</div>` : ''}
            </div>
            <div class="student-actions">
                <button class="btn btn-small btn-edit" onclick="app.showStudentModal(app.students.find(s => s.uuid === '${student.uuid}'))">编辑</button>
                <button class="btn btn-small btn-secondary" onclick="app.deleteStudent('${student.uuid}')">删除</button>
            </div>
        `;
        
        // 存储数据快照用于后续比较
        item.dataset.studentData = JSON.stringify({
            name: student.name,
            gender: student.gender,
            needsFrontSeat: student.needsFrontSeat,
            notes: student.notes
        });
    }
    
    updateClassroomContent() {
        // 增量更新教室内容（仅更新座位内容，不重建事件监听器）
        const container = document.getElementById('classroomGrid');
        
        this.seats.forEach(seat => {
            if (seat.isDeleted) return;
            
            const seatElement = container.querySelector(`[data-seat-id="${seat.id}"]`);
            if (!seatElement) return; // 如果座位不存在，跳过
            
            // 更新座位状态类
            if (seat.student) {
                seatElement.classList.remove('seat-empty');
                seatElement.classList.add('seat-occupied');
                
                // 更新性别类
                seatElement.classList.remove('male', 'female');
                if (seat.student.gender === 'male') {
                    seatElement.classList.add('male');
                } else if (seat.student.gender === 'female') {
                    seatElement.classList.add('female');
                }
                
                // 更新座位内容
                const nameLengthClass = this.getNameLengthClass(seat.student.name.length);
                const fontClass = ` font-${this.selectedFont}`;
                
                seatElement.innerHTML = `
                    <div class="seat-number">${this.rows - seat.row}-${seat.col + 1}</div>
                    <div class="student-name-display${nameLengthClass}${fontClass}" draggable="true" data-student-uuid="${seat.student.uuid}" data-source-seat-id="${seat.id}">${seat.student.name}</div>
                    <div class="seat-remove-btn" data-seat-id="${seat.id}" title="移除学生">×</div>
                `;
            } else {
                seatElement.classList.remove('seat-occupied', 'male', 'female');
                seatElement.classList.add('seat-empty');
                
                seatElement.innerHTML = `
                    <div class="seat-number">${this.rows - seat.row}-${seat.col + 1}</div>
                    <div class="seat-delete-btn" data-seat-id="${seat.id}" title="删除座位">⌫</div>
                `;
            }
        });
        
        // 重新设置移除和删除按钮的事件监听器
        this.setupSeatRemoveListeners();
        this.setupSeatDeleteListeners();
        // 重新设置学生姓名的拖拽和点击事件监听器（包括座位交换功能）
        this.setupStudentNameListeners();
    }
    
    getNameLengthClass(length) {
        if (length === 4) return ' name-4';
        if (length === 5) return ' name-5';
        if (length === 6) return ' name-6';
        if (length === 7) return ' name-7';
        if (length >= 8) return ' name-8-plus';
        return '';
    }

    renderClassroom(fullRebuild = true) {
        const container = document.getElementById('classroomGrid');
        
        // 如果不是完全重建，尝试使用增量更新
        if (!fullRebuild && container.children.length > 0) {
            this.updateClassroomContent();
            return;
        }
        
        // 完全重建模式：清空并重新创建所有元素
        container.innerHTML = '';
        container.style.gridTemplateColumns = `repeat(${this.cols}, 1fr)`;
        container.style.gridTemplateRows = `repeat(${this.rows}, 1fr) auto`;
        
        // 确保容器可以接收点击事件（添加一个透明背景区域）
        container.style.minHeight = '480px';
        container.style.height = 'auto';

        // 使用DocumentFragment批量添加元素，减少DOM重排
        const fragment = document.createDocumentFragment();

        this.seats.forEach(seat => {
            // 跳过已删除的座位，不进行渲染
            if (seat.isDeleted) {
                return;
            }
            
            const seatElement = document.createElement('div');
            seatElement.className = 'seat';
            seatElement.dataset.seatId = seat.id;
            
            // 设置明确的网格位置，确保座位保持在固定位置
            // CSS Grid 使用 1-based 索引
            seatElement.style.gridRow = seat.row + 1;
            seatElement.style.gridColumn = seat.col + 1;
            
            if (seat.student) {
                seatElement.classList.add('seat-occupied');
                if (seat.student.gender === 'male') {
                    seatElement.classList.add('male');
                } else if (seat.student.gender === 'female') {
                    seatElement.classList.add('female');
                }
                
                // 根据姓名长度确定字体大小类名
                let nameLengthClass = '';
                const nameLength = seat.student.name.length;
                if (nameLength === 4) {
                    nameLengthClass = ' name-4';
                } else if (nameLength === 5) {
                    nameLengthClass = ' name-5';
                } else if (nameLength === 6) {
                    nameLengthClass = ' name-6';
                } else if (nameLength === 7) {
                    nameLengthClass = ' name-7';
                } else if (nameLength >= 8) {
                    nameLengthClass = ' name-8-plus';
                }
                
                // 根据选择的字体添加字体类名
                let fontClass = ` font-${this.selectedFont}`;
                
                // 座位编号显示: 转换内部坐标为教室视角编号
                // this.rows - seat.row: 将内部行号转换为教室行号(底部=1, 顶部=最大)
                // seat.col + 1: 将内部列号转换为教室列号(左=1, 右=最大)
                seatElement.innerHTML = `
                    <div class="seat-number">${this.rows - seat.row}-${seat.col + 1}</div>
                    <div class="student-name-display${nameLengthClass}${fontClass}" draggable="true" data-student-uuid="${seat.student.uuid}" data-source-seat-id="${seat.id}">${seat.student.name}</div>
                    <div class="seat-remove-btn" data-seat-id="${seat.id}" title="移除学生">×</div>
                `;
            } else {
                seatElement.classList.add('seat-empty');
                // 空座位也显示教室视角的座位编号和删除按钮
                seatElement.innerHTML = `
                    <div class="seat-number">${this.rows - seat.row}-${seat.col + 1}</div>
                    <div class="seat-delete-btn" data-seat-id="${seat.id}" title="删除座位">⌫</div>
                `;
            }

            seatElement.addEventListener('click', (e) => {
                // 如果正在框选，不处理点击事件
                if (this.isSelecting) {
                    e.stopPropagation();
                    return;
                }
                
                // 只有点击座位本身或座位号时才触发单选
                if (e.target === seatElement || e.target.classList.contains('seat-number')) {
                    this.toggleSeatSelection(seat.id, false);
                    e.stopPropagation();
                }
            });
            
            // 座位鼠标按下事件处理
            seatElement.addEventListener('mousedown', (e) => {
                // 如果座位是选中状态，允许拖拽操作
                if (this.selectedSeats.has(seat.id)) {
                    // 对于选中的座位，不阻止任何事件，确保拖拽能正常启动
                    return;
                }
                
                // 对于未选中的座位，只有点击特定交互元素时才阻止框选
                if (e.target.classList.contains('student-name-display') || 
                    e.target.classList.contains('seat-remove-btn')) {
                    e.stopPropagation();
                }
            });

            // 为座位设置统一拖拽能力
            this.setupSeatDragListeners(seatElement);

            seatElement.addEventListener('dragover', (e) => {
                e.preventDefault();
                seatElement.classList.add('drag-over');
                
                // 检查是否是多选拖拽模式（只要有选中的座位就显示预览）
                if (this.selectedSeats.size > 0) {
                    // 多选拖拽：检查目标区域是否可以放置（支持位置对调）
                    const selectedSeatIds = Array.from(this.selectedSeats);
                    const dropResult = this.checkMultiDropTarget(seat.id, selectedSeatIds);
                    
                    if (dropResult) {
                        // 显示整体移动预览
                        this.showMultiDragPreview(seat.id, selectedSeatIds);
                        
                        // 检查是否会发生位置对调
                        if (dropResult.displacedStudents.length > 0) {
                            seatElement.classList.add('seat-drop-swap');
                            seatElement.classList.remove('seat-drop-target', 'seat-drop-invalid');
                        } else {
                            seatElement.classList.add('seat-drop-target');
                            seatElement.classList.remove('seat-drop-swap', 'seat-drop-invalid');
                        }
                    } else {
                        // 清除预览并显示无效状态
                        this.clearMultiDragPreview();
                        seatElement.classList.add('seat-drop-invalid');
                        seatElement.classList.remove('seat-drop-target', 'seat-drop-swap');
                    }
                } else {
                    // 单个拖拽：正常处理
                    seatElement.classList.add('seat-drop-target');
                }
            });

            seatElement.addEventListener('dragleave', () => {
                seatElement.classList.remove('drag-over', 'seat-drop-target', 'seat-drop-invalid', 'seat-drop-swap');
                // 清除拖拽预览
                this.clearMultiDragPreview();
            });

            seatElement.addEventListener('drop', (e) => {
                e.preventDefault();
                seatElement.classList.remove('drag-over', 'seat-drop-target', 'seat-drop-invalid', 'seat-drop-swap');
                // 清除拖拽预览
                this.clearMultiDragPreview();
                const dragDataString = e.dataTransfer.getData('text/plain');
                
                try {
                    const dragData = JSON.parse(dragDataString);
                    
                    // 检查是否是多选拖拽
                    if (dragData.type === 'multipleSeats') {
                        // 使用完整的选中座位信息而不是只有学生的座位
                        this.executeMultiDropWithAllSeats(seat.id, dragData);
                        return;
                    }
                    
                    // 检查是否是新的单座位拖拽格式
                    if (dragData.type === 'singleSeat') {
                        this.assignStudentToSeat(dragData.studentUuid, seat.id, dragData.sourceSeatId);
                        return;
                    }
                    
                    // 兼容旧的单个学生拖拽格式（座位间拖拽）
                    if (dragData.studentUuid && dragData.sourceSeatId) {
                        this.assignStudentToSeat(dragData.studentUuid, seat.id, dragData.sourceSeatId);
                        return;
                    }
                } catch {
                    // 从学生列表拖拽（旧格式）
                    this.assignStudentToSeat(dragDataString, seat.id, null);
                }
            });

            fragment.appendChild(seatElement);
        });

        // Add podium below the seats
        const podiumElement = document.createElement('div');
        podiumElement.className = 'podium-in-grid';
        podiumElement.style.gridColumn = `1 / -1`;
        podiumElement.style.gridRow = `${this.rows + 1}`;
        podiumElement.innerHTML = `
            <div class="podium-shape">
                <span class="podium-text">讲台</span>
            </div>
        `;
        fragment.appendChild(podiumElement);
        
        // 一次性添加所有元素到DOM，减少重排次数
        container.appendChild(fragment);

        // 为移除按钮添加事件监听器
        this.setupSeatRemoveListeners();
        
        // 为删除座位按钮添加事件监听器
        this.setupSeatDeleteListeners();
        
        // 为学生姓名添加拖拽和点击事件监听器（包括座位交换功能）
        this.setupStudentNameListeners();
        
        this.updateClassroomInfo();
        
        // 应用坐标显示设置（强制应用）
        container.classList.remove('hide-coordinates');
        if (!this.showCoordinates) {
            container.classList.add('hide-coordinates');
        }
        
    }

    setupSeatRemoveListeners() {
        document.querySelectorAll('.seat-remove-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                e.stopPropagation(); // 阻止事件冒泡到座位点击事件
                const seatId = e.target.dataset.seatId;
                this.removeStudentFromSeat(seatId);
            });
        });
    }

    setupSeatDeleteListeners() {
        document.querySelectorAll('.seat-delete-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                e.stopPropagation(); // 阻止事件冒泡到座位点击事件
                const seatId = e.target.dataset.seatId;
                if (confirm('确定要删除这个座位吗？删除后需重置才可恢复。')) {
                    this.deleteSeat(seatId);
                }
            });
        });
    }

    setupStudentNameListeners() {
        document.querySelectorAll('.student-name-display[draggable="true"]').forEach(nameElement => {
            // 标准拖拽事件（桌面端）
            nameElement.addEventListener('dragstart', (e) => {
                e.stopPropagation(); // 阻止事件冒泡到座位点击事件
                const studentUuid = e.target.dataset.studentUuid;
                const sourceSeatId = e.target.dataset.sourceSeatId;
                
                // 检查是否是多选拖拽
                if (this.startMultiDrag(e, sourceSeatId)) {
                    // 多选拖拽模式
                    return;
                }
                
                // 单个学生拖拽模式
                const dragData = JSON.stringify({
                    type: 'singleSeat',
                    studentUuid: studentUuid,
                    source: 'seat',
                    sourceSeatId: sourceSeatId
                });
                e.dataTransfer.setData('text/plain', dragData);
                
                // 添加拖拽视觉反馈
                e.target.classList.add('dragging');
                setTimeout(() => {
                    const seatElement = e.target.closest('.seat');
                    if (seatElement) {
                        seatElement.classList.add('seat-dragging');
                    }
                }, 0);
            });

            nameElement.addEventListener('dragend', (e) => {
                // 移除拖拽视觉反馈
                e.target.classList.remove('dragging');
                const seatElement = e.target.closest('.seat');
                if (seatElement) {
                    seatElement.classList.remove('seat-dragging');
                }
                // 清除拖拽预览
                this.clearMultiDragPreview();
                this.endMultiDrag(); // 清理多选拖拽状态
            });
            
            // 触摸事件（平板端）- 复用座位的触摸拖拽处理器
            const seatElement = nameElement.closest('.seat');
            if (seatElement) {
                nameElement.addEventListener('touchstart', (e) => {
                    e.stopPropagation(); // 防止触发框选
                    this.handleSeatTouchStart.call(this, {
                        ...e,
                        currentTarget: seatElement
                    });
                }, { passive: false });
            }
        });
    }

    removeStudentFromSeat(seatId) {
        const seat = this.seats.find(s => s.id === seatId);
        if (seat && seat.student) {
            // 添加到历史记录
            this.addToHistory('seatArrangement', { seats: this.seats });
            
            // 移除学生
            seat.student = null;
            
            // 更新界面（使用增量更新提升性能）
            this.saveData();
            this.renderClassroom(false);
            this.renderStudentList();
            this.updateStats();
            this.applyCurrentFilter();
        }
    }

    assignStudentToSeat(studentUuid, seatId, sourceSeatId = null) {
        const student = this.students.find(s => s.uuid === studentUuid);
        const targetSeat = this.seats.find(s => s.id === seatId);
        
        if (!student || !targetSeat || targetSeat.isDeleted) return;

        this.addToHistory('seatArrangement', { seats: this.seats });

        // 如果提供了源座位ID，优先使用它（座位间拖拽）
        const currentSeat = sourceSeatId 
            ? this.seats.find(s => s.id === sourceSeatId)
            : this.seats.find(s => s.student && s.student.uuid === studentUuid);

        // 保存目标座位上的学生（如果有）
        const displacedStudent = targetSeat.student;

        // 将拖拽的学生分配到目标座位
        targetSeat.student = student;

        // 处理被替换的学生
        if (displacedStudent) {
            // 如果有源座位，将被替换的学生放到源座位（实现交换）
            if (currentSeat && sourceSeatId) {
                currentSeat.student = displacedStudent;
            } 
            // 如果没有源座位（从学生列表拖拽），被替换的学生变为未安排状态
            else if (currentSeat) {
                currentSeat.student = displacedStudent;
            }
            // 如果拖拽的学生之前没有座位，被替换的学生将变为未安排状态（自动处理）
        } else {
            // 目标座位为空，清空源座位
            if (currentSeat) {
                currentSeat.student = null;
            }
        }
        
        this.saveData();
        this.renderClassroom(false); // 使用增量更新
        this.renderStudentList();
        this.updateStats();
        this.applyCurrentFilter();
    }

    randomSeatArrangement() {
        if (this.students.length === 0) {
            alert('请先添加学生');
            return;
        }

        this.addToHistory('seatArrangement', { seats: this.seats });
        

        // 清空所有座位（除了已删除的座位）
        this.seats.forEach(seat => {
            if (!seat.isDeleted) {
                seat.student = null;
            }
        });

        // 只选择未删除的座位作为可用座位
        const availableSeats = this.seats.filter(seat => !seat.isDeleted);
        const studentsToSeat = [...this.students];

        // 完全随机分配座位
        while (studentsToSeat.length > 0 && availableSeats.length > 0) {
            const randomStudentIndex = Math.floor(Math.random() * studentsToSeat.length);
            const randomSeatIndex = Math.floor(Math.random() * availableSeats.length);
            
            const student = studentsToSeat.splice(randomStudentIndex, 1)[0];
            const seat = availableSeats.splice(randomSeatIndex, 1)[0];
            
            seat.student = student;
        }

        this.saveData();
        this.renderClassroom(false); // 使用增量更新
        this.renderStudentList();
        this.updateStats();
        this.applyCurrentFilter();
    }

    // 根据规则排座
    ruleBasedSeatArrangement() {
        if (this.students.length === 0) {
            alert('请先添加学生');
            return;
        }

        this.addToHistory('seatArrangement', { seats: this.seats });
        

        // 清空所有座位（除了已删除的座位）
        this.seats.forEach(seat => {
            if (!seat.isDeleted) {
                seat.student = null;
            }
        });

        // 获取排座规则设置
        const arrangeByRow = document.getElementById('arrangeByRow')?.checked || false;
        const arrangeByColumn = document.getElementById('arrangeByColumn')?.checked || false;
        const heightRule = document.getElementById('heightRule')?.checked || false;
        const sameGenderRule = document.getElementById('sameGenderRule')?.checked || false;

        let studentsToSeat = [...this.students];

        // 按学号大小排序（学号越小坐得越靠前、越靠左）
        const hasStudentIdRule = arrangeByRow || arrangeByColumn;
        if (hasStudentIdRule) {
            studentsToSeat.sort((a, b) => {
                const idA = a.id || '';
                const idB = b.id || '';
                // 数字和字母混合的学号排序，学号最小的坐第一排第一列（靠近讲台的左侧）
                return idA.localeCompare(idB, undefined, { numeric: true });
            });
        }

        // 按身高排序（矮的坐前排，便于后排学生视线）
        if (heightRule) {
            studentsToSeat.sort((a, b) => {
                const heightA = parseInt(a.height) || 0;
                const heightB = parseInt(b.height) || 0;
                return heightA - heightB; // 身高矮的坐前排，身高高的坐后排
            });
        }

        // 同性别做同桌的处理
        if (sameGenderRule) {
            this.arrangeSameGenderSeating(studentsToSeat);
        } else {
            // 确定排列方式
            let arrangementType = 'row'; // 默认按行
            if (arrangeByColumn) {
                arrangementType = 'column';
            } else if (arrangeByRow) {
                arrangementType = 'row';
            }
            
            // 如果没有选择任何规则，执行随机排座
            if (!hasStudentIdRule && !heightRule && !sameGenderRule) {
                this.randomSeatArrangement();
                return; // 直接返回，不执行后续的保存和渲染（randomSeatArrangement已经包含了）
            }
            
            // 按排逐个安排
            this.arrangeStudentsInOrder(studentsToSeat, arrangementType);
        }

        this.saveData();
        this.renderClassroom(false); // 使用增量更新
        this.renderStudentList();
        this.updateStats();
        this.applyCurrentFilter();
        
        // 关闭模态框
        this.hideSeatingSettingsModal();
    }



    // 按顺序安排学生座位（遵循讲台坐标系统规则）
    arrangeStudentsInOrder(students, arrangementType = 'row') {
        // 座位排列顺序说明（重要：讲台在最下方）：
        // 1. 讲台位于教室最下方，第一排最靠近讲台（视觉上在最下方）
        // 2. 第一列位于从讲台看向学生的最左侧  
        // 3. 按行排列：第一排（最下方）从左到右，然后第二排从左到右，依此类推
        // 4. 按列排列：第一列（最左侧）从第一排开始向后排，然后第二列从第一排开始，依此类推
        // 5. 内部坐标：row值越大越靠近讲台（第一排），col值越小越靠左（第一列）
        // 6. 按学号排座时：学号越小坐得越靠前（靠近讲台）、越靠左
        
        let availableSeats;
        
        if (arrangementType === 'column') {
            // 按列排列：先按列排序，再按行排序
            // 每列从第一排（最靠近讲台，内部row最大）开始排
            availableSeats = this.seats.filter(seat => !seat.isDeleted).sort((a, b) => {
                if (a.col !== b.col) return a.col - b.col; // 先按列排序：col=0(第一列)在前
                return b.row - a.row; // 同列内按行排序：row值大的在前（第一排靠近讲台）
            });
        } else {
            // 按行排列：先按行排序，再按列排序
            // 从第一排（最靠近讲台，内部row最大）开始，从左到右排
            availableSeats = this.seats.filter(seat => !seat.isDeleted).sort((a, b) => {
                if (a.row !== b.row) return b.row - a.row; // 先按行排序：row值大的在前（第一排靠近讲台）
                return a.col - b.col; // 同行内按列排序：col=0(第一列)在前
            });
        }

        students.forEach((student, index) => {
            if (index < availableSeats.length) {
                availableSeats[index].student = student;
            }
        });
    }

    // 同性别做同桌的安排
    arrangeSameGenderSeating(students) {
        // 按性别分组
        const maleStudents = students.filter(s => s.gender === 'male');
        const femaleStudents = students.filter(s => s.gender === 'female');
        const unknownGenderStudents = students.filter(s => !s.gender || (s.gender !== 'male' && s.gender !== 'female'));

        // 获取座位，按行列排序（只选择未删除的座位）
        const availableSeats = this.seats.filter(seat => !seat.isDeleted).sort((a, b) => {
            if (a.row !== b.row) return a.row - b.row;
            return a.col - b.col;
        });

        let seatIndex = 0;

        // 优先安排男生（两个相邻座位）
        for (let i = 0; i < maleStudents.length && seatIndex < availableSeats.length; i += 2) {
            // 找到同一排的相邻两个座位
            const currentRow = Math.floor(seatIndex / this.cols);
            const seatsInRow = availableSeats.filter(seat => seat.row === currentRow);
            
            if (i + 1 < maleStudents.length && seatIndex + 1 < availableSeats.length) {
                // 安排两个男生做同桌
                availableSeats[seatIndex].student = maleStudents[i];
                availableSeats[seatIndex + 1].student = maleStudents[i + 1];
                seatIndex += 2;
            } else {
                // 只剩一个男生
                availableSeats[seatIndex].student = maleStudents[i];
                seatIndex += 1;
            }
        }

        // 安排女生（两个相邻座位）
        for (let i = 0; i < femaleStudents.length && seatIndex < availableSeats.length; i += 2) {
            if (i + 1 < femaleStudents.length && seatIndex + 1 < availableSeats.length) {
                // 安排两个女生做同桌
                availableSeats[seatIndex].student = femaleStudents[i];
                availableSeats[seatIndex + 1].student = femaleStudents[i + 1];
                seatIndex += 2;
            } else {
                // 只剩一个女生
                availableSeats[seatIndex].student = femaleStudents[i];
                seatIndex += 1;
            }
        }

        // 安排未知性别的学生
        unknownGenderStudents.forEach(student => {
            if (seatIndex < availableSeats.length) {
                availableSeats[seatIndex].student = student;
                seatIndex++;
            }
        });
    }

    clearAllSeats() {
        if (confirm('确定要重置所有座位吗？这将清空座位上的学生并恢复所有已删除的座位。')) {
            this.addToHistory('seatArrangement', { seats: this.seats });
            
            // 清空学生并恢复已删除的座位
            this.seats.forEach(seat => {
                seat.student = null;
                seat.isDeleted = false; // 恢复已删除的座位
            });
            
            this.saveData();
            this.renderClassroom(); // 使用完全重建以显示恢复的座位
            this.renderStudentList();
            this.updateStats();
            this.applyCurrentFilter(); // 应用过滤器以显示未安排的学生
        }
    }

    deleteSeat(seatId) {
        const seat = this.seats.find(s => s.id === seatId);
        if (!seat) return;
        
        // 如果座位上有学生，不允许删除
        if (seat.student) {
            alert('请先移除学生，然后才能删除座位');
            return;
        }
        
        // 记录历史状态
        this.addToHistory('seatArrangement', { seats: this.seats });
        
        // 标记座位为已删除
        seat.isDeleted = true;
        
        // 更新状态
        this.saveData();
        this.renderClassroom();
        this.renderStudentList();
        this.updateStats();
        this.applyCurrentFilter();
    }

    restoreSeat(seatId) {
        const seat = this.seats.find(s => s.id === seatId);
        if (!seat) return;
        
        // 记录历史状态
        this.addToHistory('seatArrangement', { seats: this.seats });
        
        // 恢复座位
        seat.isDeleted = false;
        
        // 更新状态
        this.saveData();
        this.renderClassroom();
        this.renderStudentList();
        this.updateStats();
        this.applyCurrentFilter();
    }

    applyNewLayout() {
        const newRows = parseInt(document.getElementById('rowCount').value);
        const newCols = parseInt(document.getElementById('colCount').value);
        
        if (newRows < 1 || newRows > 15 || newCols < 1 || newCols > 12) {
            alert('行数范围: 1-15，列数范围: 1-12');
            return;
        }

        if (confirm('改变布局将清空现有座位安排，确定继续吗？')) {
            // 记录布局改变前的状态
            this.addToHistory('seatArrangement', { seats: this.seats });
            
            this.rows = newRows;
            this.cols = newCols;
            this.initializeSeats();
            this.saveData();
            this.renderClassroom();
            this.renderStudentList();
            this.updateStats();
        }
    }

    applyNewLayoutFromDropdown() {
        const newRows = parseInt(document.getElementById('rowCountDropdown').value);
        const newCols = parseInt(document.getElementById('colCountDropdown').value);
        
        if (newRows < 1 || newRows > 15 || newCols < 1 || newCols > 12) {
            alert('行数范围: 1-15，列数范围: 1-12');
            return;
        }

        if (confirm('改变布局将清空现有座位安排，确定继续吗？')) {
            // 记录布局改变前的状态
            this.addToHistory('seatArrangement', { seats: this.seats });
            
            this.rows = newRows;
            this.cols = newCols;
            this.initializeSeats();
            this.saveData();
            this.renderClassroom();
            this.renderStudentList();
            this.updateStats();
            
            // 更新教室大小显示
            this.updateClassroomInfo();
            
            // 关闭下拉菜单
            document.getElementById('layoutSettingsDropdown').style.display = 'none';
        }
    }

    saveCurrentLayout() {
        try {
            // 创建增强版Excel数据：包含完整学生信息
            const excelData = [
                ['座位坐标', '学生姓名', '学号', '性别', '身高(cm)', '视力情况', '备注'] // 完整表头
            ];
            
            // 按列优先顺序遍历：先第1列所有座位（从后到前），再第2列所有座位，以此类推
            // 期望输出：6-1, 5-1, 4-1, 3-1, 2-1, 1-1, [空行], 6-2, 5-2, 4-2, 3-2, 2-2, 1-2, [空行]...
            // 这样符合教师视角：看教室时先看到后排，再看到前排
            for (let displayCol = 1; displayCol <= this.cols; displayCol++) {
                // 按显示行顺序从最大行数到1遍历（后排到前排，符合教师视角）
                for (let displayRow = this.rows; displayRow >= 1; displayRow--) {
                    // 将显示坐标转换为内部坐标
                    const internalRow = this.rows - displayRow;  // 显示行1 -> 内部行(this.rows-1)
                    const internalCol = displayCol - 1;          // 显示列1 -> 内部列0
                    
                    const seat = this.seats.find(s => s.row === internalRow && s.col === internalCol);
                    if (seat && !seat.isDeleted) {
                        const displayCoord = `${displayRow}-${displayCol}`;
                        
                        if (seat.student) {
                            // 有学生的座位：导出完整信息
                            const student = seat.student;
                            const genderText = student.gender === 'male' ? '男' : 
                                             student.gender === 'female' ? '女' : '';
                            const visionText = student.needsFrontSeat ? '需要前排' : '正常';
                            const heightText = student.height ? student.height.toString() : '';
                            
                            excelData.push([
                                displayCoord,           // 座位坐标
                                student.name || '',     // 学生姓名
                                student.id || '',       // 学号
                                genderText,             // 性别
                                heightText,             // 身高
                                visionText,             // 视力情况
                                student.notes || ''     // 备注
                            ]);
                        } else {
                            // 空座位：只填写坐标
                            excelData.push([displayCoord, '', '', '', '', '', '']);
                        }
                    }
                }
                
                // 在每列之后添加空行（除了最后一列）
                if (displayCol < this.cols) {
                    excelData.push(['', '', '', '', '', '', '']); // 空行
                }
            }

            // 创建工作簿
            const wb = XLSX.utils.book_new();
            const ws = XLSX.utils.aoa_to_sheet(excelData);
            
            // 设置列宽
            ws['!cols'] = [
                { width: 12 }, // 座位坐标
                { width: 12 }, // 学生姓名
                { width: 10 }, // 学号
                { width: 8 },  // 性别
                { width: 10 }, // 身高
                { width: 12 }, // 视力情况
                { width: 20 }  // 备注
            ];

            // 设置表头样式
            const headerRange = XLSX.utils.decode_range(ws['!ref']);
            for (let col = headerRange.s.c; col <= headerRange.e.c; col++) {
                const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
                if (!ws[cellAddress]) continue;
                ws[cellAddress].s = {
                    font: { bold: true },
                    fill: { fgColor: { rgb: "EFEFEF" } }
                };
            }

            XLSX.utils.book_append_sheet(wb, ws, '排座宝');
            
            // 设置工作簿属性
            wb.Props = {
                Title: "排座宝-完整版",
                Subject: "座位安排及学生信息",
                Author: "智能座位安排系统"
            };
            
            // 生成文件名（包含日期）
            const date = new Date().toLocaleDateString().replace(/\//g, '-');
            const filename = `排座宝_完整版_${date}.xlsx`;
            
            // 下载文件
            XLSX.writeFile(wb, filename, {
                bookType: 'xlsx',
                bookSST: false,
                type: 'binary'
            });
            
            alert('排座宝已导出为Excel文件！\n包含完整学生信息，方便下次导入。');
            console.log('Excel导出成功，文件名:', filename);
            console.log('导出数据包含: 座位坐标、姓名、学号、性别、身高、视力情况、备注');
            
        } catch (error) {
            console.error('Excel导出失败:', error);
            alert('导出失败，请重试');
        }
    }

    printLayout() {
        // A4纸打印功能
        // CSS打印样式已经配置，直接调用浏览器打印对话框
        
        // 先清除任何选中状态，确保打印时没有高亮
        this.clearSelection();
        
        // 直接调用浏览器打印功能（不显示提示）
        window.print();
    }

    updateStats() {
        const totalStudents = this.students.length;
        const seatedStudents = this.seats.filter(seat => seat.student).length;
        const unseatedStudents = totalStudents - seatedStudents;
        
        document.getElementById('totalStudents').textContent = totalStudents;
        document.getElementById('seatedStudents').textContent = seatedStudents;
        document.getElementById('unseatedStudents').textContent = unseatedStudents;
    }

    updateClassroomInfo() {
        document.getElementById('classroomSize').textContent = `${this.rows}行 × ${this.cols}列`;
        
    }

    filterStudents(searchTerm) {
        // 不直接设置display，而是添加/移除CSS类
        const items = document.querySelectorAll('.student-item');
        items.forEach(item => {
            const name = item.querySelector('.student-name').textContent.toLowerCase();
            const details = item.querySelector('.student-details').textContent.toLowerCase();
            
            if (name.includes(searchTerm.toLowerCase()) || details.includes(searchTerm.toLowerCase())) {
                item.classList.remove('search-hidden');
            } else {
                item.classList.add('search-hidden');
            }
        });
    }

    filterStudentsByStatus(status) {
        // 不直接设置display，而是添加/移除CSS类
        const items = document.querySelectorAll('.student-item');
        items.forEach(item => {
            const isSeated = item.classList.contains('seated');
            
            switch (status) {
                case 'all':
                    item.classList.remove('status-hidden');
                    break;
                case 'seated':
                    if (isSeated) {
                        item.classList.remove('status-hidden');
                    } else {
                        item.classList.add('status-hidden');
                    }
                    break;
                case 'unseated':
                    if (!isSeated) {
                        item.classList.remove('status-hidden');
                    } else {
                        item.classList.add('status-hidden');
                    }
                    break;
            }
        });
    }

    applyCurrentFilter() {
        const filterSelect = document.getElementById('filterStudents');
        const searchInput = document.getElementById('searchStudent');
        
        if (filterSelect && searchInput) {
            // 先清除所有筛选类
            const items = document.querySelectorAll('.student-item');
            items.forEach(item => {
                item.classList.remove('search-hidden', 'status-hidden');
            });
            
            // 应用状态筛选
            this.filterStudentsByStatus(filterSelect.value);
            
            // 应用搜索筛选
            const searchValue = searchInput.value.trim();
            if (searchValue) {
                this.filterStudents(searchValue);
            }
        }
    }

    // Excel导入相关方法
    importExcelFile() {
        const fileInput = document.getElementById('excelFileInput');
        fileInput.click();
    }

    handleExcelFileSelect(event) {
        const file = event.target.files[0];
        if (!file) return;

        if (!file.name.match(/\.(xlsx|xls)$/)) {
            alert('请选择Excel文件 (.xlsx 或 .xls)');
            return;
        }

        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                this.parseExcelData(e.target.result, file.name);
            } catch (error) {
                console.error('Excel文件读取失败:', error);
                alert('Excel文件读取失败，请检查文件格式是否正确');
            }
        };
        reader.readAsArrayBuffer(file);
    }

    parseExcelData(data, filename) {
        try {
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            
            // 转换为JSON数据
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            
            if (jsonData.length < 2) {
                alert('Excel文件中没有足够的数据');
                return;
            }

            // 获取表头和数据
            const headers = jsonData[0];
            const rows = jsonData.slice(1);
            
            // 调试信息
            console.log('Excel原始数据:', jsonData);
            console.log('表头数据:', headers);
            console.log('数据行数:', rows.length);

            // 查找列索引
            const columnMap = this.mapExcelColumns(headers);
            
            if (!columnMap.name && columnMap.name !== 0) {
                // 提供更详细的错误信息
                const headersList = headers.map((h, i) => `${i}: "${h || '(空)'}"`).join(', ');
                const headersDisplay = headers.map(h => h || '(空)').join(', ');
                console.error('未找到姓名列。当前表头:', headersList);
                alert(`Excel文件必须包含"姓名"列。\n\n当前检测到的表头: ${headersDisplay}\n\n请确保第一行包含"姓名"、"名字"或"name"列。`);
                return;
            }

            // 解析并验证数据
            const { validData, errorData } = this.validateExcelData(rows, columnMap);
            
            // 显示预览模态框
            this.showExcelPreviewModal(validData, errorData, filename);
            
        } catch (error) {
            console.error('Excel解析失败:', error);
            alert('Excel文件解析失败，请检查文件格式');
        }
    }

    mapExcelColumns(headers) {
        const columnMap = {};
        
        // 调试信息：打印headers内容
        console.log('Excel表头信息:', headers);
        
        if (!headers || !Array.isArray(headers)) {
            console.error('Invalid headers:', headers);
            return columnMap;
        }
        
        headers.forEach((header, index) => {
            if (!header && header !== 0) return; // 允许数字0作为header
            
            // 更强的字符串处理，移除所有可能的隐藏字符
            const headerStr = header.toString()
                .trim()
                .replace(/[\u200B-\u200D\uFEFF]/g, '') // 移除零宽字符
                .replace(/\s+/g, '') // 移除所有空格
                .toLowerCase();
            
            console.log(`处理表头 [${index}]: "${header}" -> "${headerStr}"`);
            
            // 座位坐标列的可能名称
            if (headerStr.includes('座位') || headerStr.includes('坐标') || headerStr.includes('位置') || 
                headerStr.includes('seat') || headerStr.includes('position') || headerStr === '座位坐标') {
                columnMap.seatCoord = index;
                console.log(`找到座位坐标列: 索引 ${index}`);
            }
            // 姓名列的可能名称（增加更多匹配选项）
            if (headerStr.includes('姓名') || headerStr.includes('名字') || headerStr.includes('name') || 
                headerStr === '姓名' || headerStr === '名字' || headerStr === 'name') {
                columnMap.name = index;
                console.log(`找到姓名列: 索引 ${index}`);
            }
            // 学号列的可能名称
            if (headerStr.includes('学号') || headerStr.includes('编号') || headerStr.includes('id') || 
                headerStr.includes('number') || headerStr === '学号' || headerStr === '编号') {
                columnMap.id = index;
            }
            // 性别列的可能名称
            if (headerStr.includes('性别') || headerStr.includes('gender') || headerStr === '性别') {
                columnMap.gender = index;
            }
            // 身高列的可能名称
            if (headerStr.includes('身高') || headerStr.includes('height') || headerStr.includes('高度') || 
                headerStr === '身高' || headerStr === '身高(cm)' || headerStr === 'height') {
                columnMap.height = index;
            }
            // 视力列的可能名称
            if (headerStr.includes('视力') || headerStr.includes('近视') || headerStr.includes('眼镜') || 
                headerStr.includes('vision') || headerStr.includes('情况')) {
                columnMap.vision = index;
            }
            // 备注列的可能名称
            if (headerStr.includes('备注') || headerStr.includes('说明') || headerStr.includes('note') || 
                headerStr.includes('remark') || headerStr === '备注') {
                columnMap.notes = index;
            }
        });
        
        console.log('列映射结果:', columnMap);
        return columnMap;
    }

    validateExcelData(rows, columnMap) {
        const validData = [];
        const errorData = [];

        rows.forEach((row, rowIndex) => {
            const actualRow = rowIndex + 2; // Excel行号（从1开始，加上表头行）
            const errors = [];
            
            // 检查是否为空行
            if (!row || row.every(cell => !cell || cell.toString().trim() === '')) {
                return; // 跳过空行
            }

            const studentData = {
                name: row[columnMap.name] ? row[columnMap.name].toString().trim() : '',
                id: row[columnMap.id] ? row[columnMap.id].toString().trim() : '',
                gender: row[columnMap.gender] ? row[columnMap.gender].toString().trim() : '',
                height: row[columnMap.height] ? row[columnMap.height].toString().trim() : '',
                vision: row[columnMap.vision] ? row[columnMap.vision].toString().trim() : '',
                notes: row[columnMap.notes] ? row[columnMap.notes].toString().trim() : '',
                seatCoord: row[columnMap.seatCoord] ? row[columnMap.seatCoord].toString().trim() : ''
            };

            // 验证必填字段
            if (!studentData.name) {
                errors.push('姓名不能为空');
            }

            // 验证性别
            if (studentData.gender) {
                const genderLower = studentData.gender.toLowerCase();
                if (genderLower.includes('男') || genderLower.includes('male') || genderLower === 'm') {
                    studentData.gender = 'male';
                } else if (genderLower.includes('女') || genderLower.includes('female') || genderLower === 'f') {
                    studentData.gender = 'female';
                } else {
                    studentData.gender = '';
                }
            }

            // 验证身高
            if (studentData.height) {
                const originalHeight = studentData.height;
                const heightNum = parseInt(studentData.height);
                if (!isNaN(heightNum) && heightNum >= 100 && heightNum <= 250) {
                    studentData.height = heightNum;
                } else {
                    studentData.height = null;
                    if (originalHeight !== '') {
                        errors.push('身高应为100-250cm之间的数字');
                    }
                }
            } else {
                studentData.height = null;
            }

            // 验证视力信息
            if (studentData.vision) {
                const visionLower = studentData.vision.toLowerCase();
                if (visionLower.includes('近视') || visionLower.includes('不佳') || 
                    visionLower.includes('戴眼镜') || visionLower.includes('眼镜') ||
                    visionLower.includes('poor') || visionLower.includes('bad') ||
                    visionLower.includes('yes') || visionLower.includes('是') || visionLower.includes('需要')) {
                    studentData.needsFrontSeat = true;
                } else {
                    studentData.needsFrontSeat = false;
                }
            } else {
                studentData.needsFrontSeat = false;
            }

            if (errors.length > 0) {
                errorData.push({
                    row: actualRow,
                    data: studentData,
                    errors: errors,
                    originalData: row
                });
            } else {
                validData.push(studentData);
            }
        });

        return { validData, errorData };
    }

    showExcelPreviewModal(validData, errorData, filename) {
        this.importData = { validData, errorData, filename };
        
        // 更新统计信息
        document.getElementById('totalImportCount').textContent = validData.length + errorData.length;
        document.getElementById('validImportCount').textContent = validData.length;
        document.getElementById('errorImportCount').textContent = errorData.length;

        // 渲染有效数据表格
        this.renderValidDataTable(validData);
        
        // 渲染错误数据列表
        this.renderErrorDataList(errorData);
        
        // 显示模态框
        document.getElementById('excelPreviewModal').style.display = 'flex';
        
        // 默认显示有效数据标签页
        this.switchTab('valid');
    }

    renderValidDataTable(validData) {
        const tbody = document.getElementById('validDataBody');
        tbody.innerHTML = '';

        validData.forEach((student, index) => {
            const row = document.createElement('tr');
            row.innerHTML = `
                <td>${student.name}</td>
                <td>${student.id || '<span class="empty-cell">无</span>'}</td>
                <td>${student.gender === 'male' ? '男' : student.gender === 'female' ? '女' : '<span class="empty-cell">无</span>'}</td>
                <td>${student.height ? student.height + 'cm' : '<span class="empty-cell">无</span>'}</td>
                <td>${student.needsFrontSeat ? '需要前排' : '正常'}</td>
                <td>${student.notes || '<span class="empty-cell">无</span>'}</td>
            `;
            tbody.appendChild(row);
        });
    }

    renderErrorDataList(errorData) {
        const container = document.getElementById('errorsList');
        container.innerHTML = '';

        if (errorData.length === 0) {
            container.innerHTML = '<div style="text-align: center; color: #6b7280; padding: 2rem;">没有错误数据</div>';
            return;
        }

        errorData.forEach(error => {
            const errorItem = document.createElement('div');
            errorItem.className = 'error-item';
            
            errorItem.innerHTML = `
                <div class="error-row">第 ${error.row} 行</div>
                <div class="error-message">${error.errors.join('、')}</div>
                <div class="error-data">
                    原始数据: ${JSON.stringify(error.originalData).replace(/"/g, '')}
                </div>
            `;
            
            container.appendChild(errorItem);
        });
    }

    switchTab(tabName) {
        // 更新标签按钮状态
        document.querySelectorAll('.tab-btn').forEach(btn => {
            btn.classList.remove('active');
        });
        document.querySelector(`[data-tab="${tabName}"]`).classList.add('active');

        // 更新内容显示
        document.querySelectorAll('.tab-content').forEach(content => {
            content.classList.remove('active');
        });
        document.getElementById(`${tabName}DataPreview`).classList.add('active');
    }

    hideExcelPreviewModal() {
        document.getElementById('excelPreviewModal').style.display = 'none';
        this.importData = null;
        
        // 重置文件输入
        document.getElementById('excelFileInput').value = '';
    }

    confirmExcelImport() {
        if (!this.importData || !this.importData.validData.length) {
            alert('没有有效数据可以导入');
            return;
        }

        const overwriteExisting = document.getElementById('overwriteExisting').checked;
        const skipInvalid = document.getElementById('skipInvalid').checked;
        
        let importCount = 0;
        let updateCount = 0;
        let skipCount = 0;
        let seatAssignmentCount = 0;
        let conflictCount = 0;
        let invalidCoordCount = 0;

        // 检查是否有座位坐标需要处理，如果有则添加历史记录
        const hasSeatingData = this.importData.validData.some(student => student.seatCoord);
        if (hasSeatingData) {
            this.addToHistory('seatArrangement', { seats: this.seats });
        }

        this.importData.validData.forEach(studentData => {
            // 检查是否存在同名学生
            const existingStudent = this.students.find(s => s.name === studentData.name);
            let currentStudent = null;
            
            if (existingStudent) {
                if (overwriteExisting) {
                    // 更新现有学生信息
                    existingStudent.id = studentData.id || existingStudent.id;
                    existingStudent.gender = studentData.gender || existingStudent.gender;
                    existingStudent.height = studentData.height !== null ? studentData.height : existingStudent.height;
                    existingStudent.needsFrontSeat = studentData.needsFrontSeat;
                    existingStudent.notes = studentData.notes || existingStudent.notes;
                    currentStudent = existingStudent;
                    updateCount++;
                } else {
                    skipCount++;
                    return; // 跳过这个学生，不进行座位分配
                }
            } else {
                // 添加新学生
                const newStudent = {
                    uuid: this.generateUUID(),
                    name: studentData.name,
                    id: studentData.id,
                    gender: studentData.gender,
                    height: studentData.height,
                    needsFrontSeat: studentData.needsFrontSeat,
                    notes: studentData.notes,
                    seatId: null
                };
                this.students.push(newStudent);
                currentStudent = newStudent;
                importCount++;
            }

            // 处理座位坐标分配
            if (studentData.seatCoord && currentStudent) {
                const parsedCoord = this.parseDisplayCoordinate(studentData.seatCoord);
                
                if (parsedCoord) {
                    // 查找对应的座位
                    const targetSeat = this.seats.find(seat => 
                        seat.row === parsedCoord.row && seat.col === parsedCoord.col
                    );
                    
                    if (targetSeat) {
                        if (targetSeat.student) {
                            // 座位已被占用，记录冲突但不强制替换
                            conflictCount++;
                            console.warn(`座位 ${studentData.seatCoord} 已被 ${targetSeat.student.name} 占用，学生 ${currentStudent.name} 未分配座位`);
                        } else {
                            // 清空学生当前的座位（如果有的话）
                            const currentSeat = this.seats.find(seat => seat.student && seat.student.uuid === currentStudent.uuid);
                            if (currentSeat) {
                                currentSeat.student = null;
                            }
                            
                            // 分配到新座位
                            targetSeat.student = currentStudent;
                            seatAssignmentCount++;
                            console.log(`学生 ${currentStudent.name} 已分配到座位 ${studentData.seatCoord}`);
                        }
                    } else {
                        invalidCoordCount++;
                        console.warn(`未找到座位 ${studentData.seatCoord}，学生 ${currentStudent.name} 未分配座位`);
                    }
                } else {
                    invalidCoordCount++;
                    console.warn(`无效的座位坐标格式: ${studentData.seatCoord}，学生 ${currentStudent.name} 未分配座位`);
                }
            }
        });

        // 保存数据并更新界面
        this.saveData();
        this.renderStudentList();
        this.renderClassroom();
        this.updateStats();
        this.applyCurrentFilter();
        
        // 显示导入结果
        let message = `导入完成！\n新增学生: ${importCount} 人`;
        if (updateCount > 0) {
            message += `\n更新学生: ${updateCount} 人`;
        }
        if (skipCount > 0) {
            message += `\n跳过重复: ${skipCount} 人`;
        }
        if (seatAssignmentCount > 0) {
            message += `\n座位安排: ${seatAssignmentCount} 人`;
        }
        if (conflictCount > 0) {
            message += `\n座位冲突: ${conflictCount} 人（已跳过）`;
        }
        if (invalidCoordCount > 0) {
            message += `\n无效坐标: ${invalidCoordCount} 人（已跳过）`;
        }
        
        alert(message);
        
        // 关闭模态框
        this.hideExcelPreviewModal();
    }



    clearAllStudents() {
        const confirmMessage = `
确定要清空所有数据吗？

此操作将：
• 删除所有学生信息
• 清空所有座位安排
• 重置所有设置到初始状态
• 清空操作历史记录

⚠️ 此操作无法撤销！
        `.trim();

        if (confirm(confirmMessage)) {
            // 清空所有数据
            this.students = [];
            // 清空所有座位上的学生信息，但保留座位的删除状态
            this.seats.forEach(seat => {
                seat.student = null;
            });
            this.history = [];
            this.historyIndex = -1;
            this.constraints = [];

            // 重置筛选器到默认状态
            document.getElementById('filterStudents').value = 'unseated';
            document.getElementById('searchStudent').value = '';

            // 保存并更新界面
            this.saveData();
            this.renderClassroom();
            this.renderStudentList();
            this.updateStats();
            this.updateHistoryButtons();
            this.applyCurrentFilter();

            // 显示成功消息
            alert('所有数据已清空，系统已重置到初始状态！');
        }
    }

    showSeatingSettingsModal() {
        document.getElementById('seatingSettingsModal').style.display = 'flex';
        // 更新约束列表显示
        this.renderConstraintList();
        // 更新字体选择下拉框的值
        document.getElementById('fontSelect').value = this.selectedFont;
    }

    hideSeatingSettingsModal() {
        document.getElementById('seatingSettingsModal').style.display = 'none';
    }

    toggleCoordinatesDisplay(show) {
        this.showCoordinates = show;
        this.saveData();
        
        const classroomGrid = document.getElementById('classroomGrid');
        if (show) {
            classroomGrid.classList.remove('hide-coordinates');
        } else {
            classroomGrid.classList.add('hide-coordinates');
        }
    }

    toggleLayoutSettingsDropdown() {
        const dropdown = document.getElementById('layoutSettingsDropdown');
        if (dropdown.style.display === 'none' || dropdown.style.display === '') {
            dropdown.style.display = 'block';
        } else {
            dropdown.style.display = 'none';
        }
    }

    hideLayoutSettingsDropdown() {
        const dropdown = document.getElementById('layoutSettingsDropdown');
        dropdown.style.display = 'none';
    }

    initializeLayoutSettings() {
        const checkbox = document.getElementById('showCoordinatesToggle');
        if (checkbox) {
            checkbox.checked = this.showCoordinates;
        }
        
        // 强制应用坐标显示设置
        const classroomGrid = document.getElementById('classroomGrid');
        if (classroomGrid) {
            if (this.showCoordinates) {
                classroomGrid.classList.remove('hide-coordinates');
            } else {
                classroomGrid.classList.add('hide-coordinates');
            }
        }
        
        const fontSelect = document.getElementById('fontSelectDropdown');
        if (fontSelect) {
            fontSelect.value = this.selectedFont;
        }
        
        // Initialize layout input fields in dropdown
        const rowCountDropdown = document.getElementById('rowCountDropdown');
        const colCountDropdown = document.getElementById('colCountDropdown');
        if (rowCountDropdown) {
            rowCountDropdown.value = this.rows;
        }
        if (colCountDropdown) {
            colCountDropdown.value = this.cols;
        }
        
        // Initialize color selections in the UI
        this.initializeColorSelections();
        
        // Apply the current colors
        this.updateSeatColors();
    }

    initializeColorSelections() {
        // Update male color selection
        const maleSwatches = document.querySelectorAll('.color-swatches[data-gender="male"] .color-swatch');
        maleSwatches.forEach(swatch => {
            swatch.classList.remove('selected');
            if (swatch.dataset.color === this.maleColor) {
                swatch.classList.add('selected');
            }
        });
        
        // Update female color selection
        const femaleSwatches = document.querySelectorAll('.color-swatches[data-gender="female"] .color-swatch');
        femaleSwatches.forEach(swatch => {
            swatch.classList.remove('selected');
            if (swatch.dataset.color === this.femaleColor) {
                swatch.classList.add('selected');
            }
        });
    }

    changeFontFamily(fontValue) {
        this.selectedFont = fontValue;
        this.saveData();
        this.renderClassroom(); // 重新渲染座位以应用新字体
        this.renderStudentList(); // 同时更新学生列表
        this.updateStats();
        this.applyCurrentFilter();
    }

    selectColor(swatchElement) {
        const color = swatchElement.dataset.color;
        const colorContainer = swatchElement.closest('.color-swatches');
        if (!colorContainer) return;
        const gender = colorContainer.dataset.gender;
        
        // Update the color property
        if (gender === 'male') {
            this.maleColor = color;
        } else if (gender === 'female') {
            this.femaleColor = color;
        }
        
        // Update UI: remove selected class from siblings and add to clicked element
        const swatches = swatchElement.closest('.color-swatches');
        if (swatches) {
            swatches.querySelectorAll('.color-swatch').forEach(swatch => {
                swatch.classList.remove('selected');
            });
        }
        swatchElement.classList.add('selected');
        
        // Update seat colors immediately
        this.updateSeatColors();
        
        // Save data
        this.saveData();
    }

    updateSeatColors() {
        // Apply colors using CSS custom properties
        document.documentElement.style.setProperty('--male-color', this.maleColor);
        document.documentElement.style.setProperty('--female-color', this.femaleColor);
    }

    addConstraint() {
        const constraintInput = document.getElementById('constraintInput');
        const constraintText = constraintInput.value.trim();
        
        if (!constraintText) {
            alert('请输入约束条件');
            return;
        }
        
        // 创建约束对象
        const constraint = {
            id: this.generateUUID(),
            text: constraintText,
            type: 'custom',
            active: true,
            timestamp: Date.now()
        };
        
        // 添加到约束列表
        this.constraints.push(constraint);
        
        // 保存数据
        this.saveData();
        
        // 清空输入框
        constraintInput.value = '';
        
        // 更新约束列表显示
        this.renderConstraintList();
        
        // 显示成功消息
        console.log('约束条件已添加:', constraint);
    }

    renderConstraintList() {
        const container = document.getElementById('constraintList');
        container.innerHTML = '';
        
        if (this.constraints.length === 0) {
            container.innerHTML = '<div class="no-constraints">暂无约束条件</div>';
            return;
        }
        
        this.constraints.forEach(constraint => {
            const constraintItem = document.createElement('div');
            constraintItem.className = 'constraint-item';
            constraintItem.innerHTML = `
                <div class="constraint-text">${constraint.text}</div>
                <div class="constraint-actions">
                    <button class="btn btn-small btn-secondary" onclick="app.removeConstraint('${constraint.id}')">删除</button>
                </div>
            `;
            container.appendChild(constraintItem);
        });
    }

    removeConstraint(constraintId) {
        if (confirm('确定要删除这个约束条件吗？')) {
            this.constraints = this.constraints.filter(c => c.id !== constraintId);
            this.saveData();
            this.renderConstraintList();
        }
    }

    // 多选功能事件监听器设置
    setupMultiSelectEventListeners() {
        const classroomGrid = document.getElementById('classroomGrid');
        
        // 工具栏按钮事件（添加null检查）
        const selectAllBtn = document.getElementById('selectAllSeats');
        const clearSelectionBtn = document.getElementById('clearSelection');
        const clearSeatsBtn = document.getElementById('clearSelectedSeats');
        
        if (selectAllBtn) {
            selectAllBtn.addEventListener('click', () => {
                this.selectAllSeats();
            });
        }
        
        if (clearSelectionBtn) {
            clearSelectionBtn.addEventListener('click', () => {
                this.clearSelection();
            });
        }
        
        if (clearSeatsBtn) {
            clearSeatsBtn.addEventListener('click', () => {
                this.clearSelectedSeats();
            });
        }

        // 统一的事件处理函数
        const handleStart = (e) => {
            const clickedElement = e.target;
            
            // 检查是否点击在选中的座位上
            const clickedSeat = clickedElement.closest('.seat');
            if (clickedSeat) {
                const seatId = clickedSeat.dataset.seatId;
                if (this.selectedSeats.has(seatId)) {
                    return; // 不启动框选，让拖拽事件正常处理
                }
            }
            
            // 只排除这些特定的交互元素，其他所有地方都可以框选
            const shouldExclude = 
                clickedElement.classList.contains('student-name-display') ||
                clickedElement.classList.contains('seat-remove-btn') ||
                clickedElement.classList.contains('multi-select-btn') ||
                clickedElement.closest('.multi-select-toolbar') !== null;
            
            // 在空白区域开始框选
            if (!shouldExclude) {
                this.startSelection(e);
            }
        };

        const handleMove = (e) => {
            if (this.isSelecting) {
                this.updateSelection(e);
            }
        };

        const handleEnd = (e) => {
            if (this.isSelecting) {
                this.endSelection(e);
            }
        };

        // 鼠标事件（桌面端）
        classroomGrid.addEventListener('mousedown', handleStart);
        classroomGrid.addEventListener('mousemove', handleMove);
        classroomGrid.addEventListener('mouseup', handleEnd);

        // 触摸事件（平板端）
        classroomGrid.addEventListener('touchstart', handleStart, { passive: false });
        classroomGrid.addEventListener('touchmove', handleMove, { passive: false });
        classroomGrid.addEventListener('touchend', handleEnd, { passive: false });

        // 全局键盘事件
        document.addEventListener('keydown', (e) => {
            if (e.key === 'Escape') {
                this.clearSelection();
            }
            if (e.key === 'Delete' && this.selectedSeats.size > 0) {
                this.clearSelectedSeats();
            }
        });
    }

    // 开始框选
    startSelection(e) {
        // 每次框选都先清空之前的选择，简化操作
        this.clearSelection();
        
        this.isSelecting = true;
        const rect = e.currentTarget.getBoundingClientRect();
        
        // 统一处理触摸和鼠标事件
        const clientX = e.touches ? e.touches[0].clientX : e.clientX;
        const clientY = e.touches ? e.touches[0].clientY : e.clientY;
        
        this.selectionStart = {
            x: clientX - rect.left,
            y: clientY - rect.top
        };

        // 创建选择框
        this.selectionBox = document.createElement('div');
        this.selectionBox.className = 'selection-box';
        this.selectionBox.style.left = this.selectionStart.x + 'px';
        this.selectionBox.style.top = this.selectionStart.y + 'px';
        this.selectionBox.style.width = '0px';
        this.selectionBox.style.height = '0px';
        e.currentTarget.appendChild(this.selectionBox);
        
        // 禁用滚动（触摸设备）
        e.currentTarget.classList.add('is-selecting');

        e.preventDefault();
    }

    // 更新框选
    updateSelection(e) {
        if (!this.isSelecting || !this.selectionBox) return;

        const rect = e.currentTarget.getBoundingClientRect();
        
        // 统一处理触摸和鼠标事件
        const clientX = e.touches ? e.touches[0].clientX : e.clientX;
        const clientY = e.touches ? e.touches[0].clientY : e.clientY;
        
        const currentX = clientX - rect.left;
        const currentY = clientY - rect.top;

        const left = Math.min(this.selectionStart.x, currentX);
        const top = Math.min(this.selectionStart.y, currentY);
        const width = Math.abs(currentX - this.selectionStart.x);
        const height = Math.abs(currentY - this.selectionStart.y);

        this.selectionBox.style.left = left + 'px';
        this.selectionBox.style.top = top + 'px';
        this.selectionBox.style.width = width + 'px';
        this.selectionBox.style.height = height + 'px';

        // 实时更新选中的座位
        this.updateSelectedSeatsInBox(left, top, width, height);
        
        // 防止触摸滚动
        e.preventDefault();
    }

    // 结束框选
    endSelection(e) {
        this.isSelecting = false;
        
        // 将框选中的座位加入正式选择
        let selectedCount = 0;
        document.querySelectorAll('.seat-multi-selecting').forEach(element => {
            const seatId = element.dataset.seatId;
            if (seatId) {
                this.selectedSeats.add(seatId);
                selectedCount++;
            }
            element.classList.remove('seat-multi-selecting');
        });
        
        if (this.selectionBox) {
            this.selectionBox.remove();
            this.selectionBox = null;
        }
        
        this.selectionStart = null;
        this.updateSelectionUI();
        
        // 恢复滚动（触摸设备）
        const classroomGrid = document.getElementById('classroomGrid');
        if (classroomGrid) {
            classroomGrid.classList.remove('is-selecting');
        }

    }

    // 更新框选范围内的座位
    updateSelectedSeatsInBox(left, top, width, height) {
        const seatElements = document.querySelectorAll('.seat');
        const tempSelected = new Set();

        seatElements.forEach(seatElement => {
            const seatRect = seatElement.getBoundingClientRect();
            const gridContainer = seatElement.closest('.classroom-grid');
            if (!gridContainer) return;
            const gridRect = gridContainer.getBoundingClientRect();
            
            // 获取座位相对于教室网格的位置
            const seatLeft = seatRect.left - gridRect.left;
            const seatTop = seatRect.top - gridRect.top;
            const seatRight = seatLeft + seatRect.width;
            const seatBottom = seatTop + seatRect.height;
            
            const selectionRight = left + width;
            const selectionBottom = top + height;

            // 改进碰撞检测：要求选择框与座位有实际重叠（不仅仅是边界接触）
            const hasOverlap = (
                seatLeft < selectionRight && 
                seatRight > left && 
                seatTop < selectionBottom && 
                seatBottom > top
            );

            // 更严格的选择条件：座位中心点必须在选择框内，或者有足够的重叠面积
            const seatCenterX = seatLeft + seatRect.width / 2;
            const seatCenterY = seatTop + seatRect.height / 2;
            const centerInSelection = (
                seatCenterX >= left && seatCenterX <= selectionRight &&
                seatCenterY >= top && seatCenterY <= selectionBottom
            );

            // 计算重叠面积（作为备选条件）
            const overlapLeft = Math.max(seatLeft, left);
            const overlapTop = Math.max(seatTop, top);
            const overlapRight = Math.min(seatRight, selectionRight);
            const overlapBottom = Math.min(seatBottom, selectionBottom);
            const overlapArea = Math.max(0, overlapRight - overlapLeft) * Math.max(0, overlapBottom - overlapTop);
            const seatArea = seatRect.width * seatRect.height;
            const overlapRatio = overlapArea / seatArea;

            // 选择条件：中心点在框内 或 重叠面积超过30%
            if (hasOverlap && (centerInSelection || overlapRatio > 0.3)) {
                const seatId = seatElement.dataset.seatId;
                if (seatId) {
                    // 检查座位是否已删除，如果已删除则跳过选择
                    const seat = this.seats.find(s => s.id === seatId);
                    if (seat && !seat.isDeleted) {
                        tempSelected.add(seatId);
                    }
                }
            }
        });

        // 清除当前框选状态
        document.querySelectorAll('.seat-multi-selecting').forEach(element => {
            element.classList.remove('seat-multi-selecting');
        });

        // 添加框选状态
        tempSelected.forEach(seatId => {
            const seatElement = document.querySelector(`[data-seat-id="${seatId}"]`);
            if (seatElement) {
                seatElement.classList.add('seat-multi-selecting');
            }
        });

        // 实时更新工具栏显示（框选过程中的临时状态）
        if (tempSelected.size > 0) {
            const toolbar = document.getElementById('multiSelectToolbar');
            const countElement = document.getElementById('selectionCount');
            if (toolbar && countElement) {
                toolbar.classList.add('show');
                countElement.textContent = `正在框选 ${tempSelected.size} 个座位`;
            }
        }
    }

    // 点击座位选择（简化多选）
    toggleSeatSelection(seatId, clearOthers = false) {
        if (clearOthers && !this.selectedSeats.has(seatId)) {
            this.clearSelection();
        }

        if (this.selectedSeats.has(seatId)) {
            this.selectedSeats.delete(seatId);
        } else {
            this.selectedSeats.add(seatId);
        }

        this.updateSelectionUI();
    }

    // 全选座位
    selectAllSeats() {
        this.selectedSeats.clear();
        this.seats.forEach(seat => {
            if (!seat.isDeleted) {
                this.selectedSeats.add(seat.id);
            }
        });
        this.updateSelectionUI();
    }

    // 清空选择
    clearSelection() {
        this.selectedSeats.clear();
        this.updateSelectionUI();
    }

    // 批量清空选中座位上的学生
    clearSelectedSeats() {
        if (this.selectedSeats.size === 0) {
            alert('请先选择要清空的座位');
            return;
        }

        if (confirm(`确定要清空选中的 ${this.selectedSeats.size} 个座位吗？`)) {
            this.addToHistory('seatArrangement', { seats: this.seats });
            
            this.selectedSeats.forEach(seatId => {
                const seat = this.seats.find(s => s.id === seatId);
                if (seat && !seat.isDeleted) {
                    seat.student = null;
                }
            });

            this.clearSelection();
            this.saveData();
            this.renderClassroom(false); // 使用增量更新
            this.renderStudentList();
            this.updateStats();
            this.applyCurrentFilter();
        }
    }

    // 更新选择UI状态（简化版本，只负责视觉反馈）
    updateSelectionUI() {
        const toolbar = document.getElementById('multiSelectToolbar');
        const countElement = document.getElementById('selectionCount');
        
        // 更新所有座位的选中状态显示
        document.querySelectorAll('.seat').forEach(seatElement => {
            const seatId = seatElement.dataset.seatId;
            seatElement.classList.remove('seat-multi-selected', 'seat-multi-selecting');
            
            if (this.selectedSeats.has(seatId)) {
                seatElement.classList.add('seat-multi-selected');
            }
        });

        // 更新工具栏显示
        if (toolbar && countElement) {
            if (this.selectedSeats.size > 0) {
                toolbar.classList.add('show');
                countElement.textContent = `已选择 ${this.selectedSeats.size} 个座位`;
            } else {
                toolbar.classList.remove('show');
            }
        }
    }

    // 统一的座位拖拽事件处理器
    handleSeatDragStart(e) {
        const seatElement = e.currentTarget;
        const seatId = seatElement.dataset.seatId;
        const seat = this.seats.find(s => s.id === seatId);
        
        // 如果点击的是移除按钮，阻止拖拽
        if (e.target.classList.contains('seat-remove-btn')) {
            e.preventDefault();
            return;
        }
        
        // 防止重复拖拽
        if (this.isDragging || this.isMultiDragging) {
            e.preventDefault();
            return;
        }
        
        // 运行时决策：根据当前状态判断拖拽类型
        const isInMultiSelection = this.selectedSeats.has(seatId);
        const hasMultipleSelections = this.selectedSeats.size > 1;
        
        if (isInMultiSelection && hasMultipleSelections) {
            // 多座位拖拽模式
            this.startMultiSeatDrag(e, seatId);
        } else {
            // 单座位拖拽模式
            this.startSingleSeatDrag(e, seatId, seat);
        }
    }
    
    // 单座位拖拽启动
    startSingleSeatDrag(e, seatId, seat) {
        if (!seat || !seat.student) {
            e.preventDefault();
            return;
        }
        
        this.isDragging = true;
        
        const dragData = {
            type: 'singleSeat',
            studentUuid: seat.student.uuid,
            sourceSeatId: seatId
        };
        
        e.dataTransfer.setData('text/plain', JSON.stringify(dragData));
        e.currentTarget.classList.add('seat-dragging');
        e.currentTarget.style.cursor = 'grabbing';
    }
    
    // 多座位拖拽启动
    startMultiSeatDrag(e, seatId) {
        // 收集所有选中座位的数据
        const selectedSeatsData = [];
        const studentsData = [];
        
        this.selectedSeats.forEach(selectedSeatId => {
            const selectedSeat = this.seats.find(s => s.id === selectedSeatId);
            if (selectedSeat) {
                selectedSeatsData.push({
                    seatId: selectedSeatId,
                    row: selectedSeat.row,
                    col: selectedSeat.col,
                    student: selectedSeat.student
                });
                
                if (selectedSeat.student) {
                    studentsData.push(selectedSeat.student);
                }
            }
        });
        
        this.isMultiDragging = true;
        
        const dragData = {
            type: 'multipleSeats',
            sourceSeats: selectedSeatsData,
            studentsData: studentsData
        };
        
        e.dataTransfer.setData('text/plain', JSON.stringify(dragData));
        
        // 为所有选中座位添加拖拽样式
        this.selectedSeats.forEach(selectedSeatId => {
            const element = document.querySelector(`[data-seat-id="${selectedSeatId}"]`);
            if (element) element.classList.add('seat-drag-preview');
        });
        
        // 创建自定义拖拽图像
        this.createMultiDragImage(e, this.selectedSeats.size);
    }
    
    // 统一的拖拽结束处理
    handleSeatDragEnd(e) {
        // 清理所有拖拽状态
        this.isDragging = false;
        this.isMultiDragging = false;
        
        // 移除拖拽样式
        document.querySelectorAll('.seat-dragging, .seat-drag-preview')
            .forEach(el => {
                el.classList.remove('seat-dragging', 'seat-drag-preview');
                el.style.cursor = '';
            });
        
        this.clearMultiDragPreview();
    }

    // 设置座位的统一拖拽能力
    setupSeatDragListeners(seatElement) {
        // 移除旧的监听器（如果存在）
        this.removeSeatDragListeners(seatElement);
        
        // 设置基本拖拽属性
        const seat = this.seats.find(s => s.id === seatElement.dataset.seatId);
        if (seat && seat.student) {
            seatElement.draggable = true;
            seatElement.style.cursor = 'grab';
        } else {
            seatElement.draggable = false;
            seatElement.style.cursor = '';
        }
        
        // 绑定统一的事件处理器
        const boundDragStart = this.handleSeatDragStart.bind(this);
        const boundDragEnd = this.handleSeatDragEnd.bind(this);
        const boundTouchStart = this.handleSeatTouchStart.bind(this);
        const boundTouchMove = this.handleSeatTouchMove.bind(this);
        const boundTouchEnd = this.handleSeatTouchEnd.bind(this);
        
        seatElement._unifiedDragStartHandler = boundDragStart;
        seatElement._unifiedDragEndHandler = boundDragEnd;
        seatElement._unifiedTouchStartHandler = boundTouchStart;
        seatElement._unifiedTouchMoveHandler = boundTouchMove;
        seatElement._unifiedTouchEndHandler = boundTouchEnd;
        
        // 添加标准拖拽事件监听器（桌面端）
        seatElement.addEventListener('dragstart', boundDragStart);
        seatElement.addEventListener('dragend', boundDragEnd);
        
        // 添加触摸事件监听器（平板端）
        seatElement.addEventListener('touchstart', boundTouchStart, { passive: false });
        seatElement.addEventListener('touchmove', boundTouchMove, { passive: false });
        seatElement.addEventListener('touchend', boundTouchEnd, { passive: false });
    }
    
    // 移除座位的拖拽监听器
    removeSeatDragListeners(seatElement) {
        if (seatElement._unifiedDragStartHandler) {
            seatElement.removeEventListener('dragstart', seatElement._unifiedDragStartHandler);
            delete seatElement._unifiedDragStartHandler;
        }
        if (seatElement._unifiedDragEndHandler) {
            seatElement.removeEventListener('dragend', seatElement._unifiedDragEndHandler);
            delete seatElement._unifiedDragEndHandler;
        }
        if (seatElement._unifiedTouchStartHandler) {
            seatElement.removeEventListener('touchstart', seatElement._unifiedTouchStartHandler);
            delete seatElement._unifiedTouchStartHandler;
        }
        if (seatElement._unifiedTouchMoveHandler) {
            seatElement.removeEventListener('touchmove', seatElement._unifiedTouchMoveHandler);
            delete seatElement._unifiedTouchMoveHandler;
        }
        if (seatElement._unifiedTouchEndHandler) {
            seatElement.removeEventListener('touchend', seatElement._unifiedTouchEndHandler);
            delete seatElement._unifiedTouchEndHandler;
        }
    }

    // 废弃的函数已移除，使用 removeSeatDragListeners(seatElement) 替代

    // 触摸拖拽处理 - 开始
    handleSeatTouchStart(e) {
        const seatElement = e.currentTarget;
        const seatId = seatElement.dataset.seatId;
        const seat = this.seats.find(s => s.id === seatId);
        
        // 如果点击的是移除按钮，不处理拖拽
        if (e.target.classList.contains('seat-remove-btn')) {
            return;
        }
        
        // 没有学生的座位不能拖拽
        if (!seat || !seat.student) {
            return;
        }
        
        // 记录触摸开始时间和位置
        this.touchStartTime = Date.now();
        const touch = e.touches[0];
        this.touchStartX = touch.clientX;
        this.touchStartY = touch.clientY;
        
        // 准备拖拽数据
        const isInMultiSelection = this.selectedSeats.has(seatId);
        const hasMultipleSelections = this.selectedSeats.size > 1;
        
        if (isInMultiSelection && hasMultipleSelections) {
            // 多座位拖拽
            const selectedStudents = [];
            this.selectedSeats.forEach(sid => {
                const s = this.seats.find(seat => seat.id === sid);
                if (s && s.student) {
                    selectedStudents.push({
                        seatId: sid,
                        student: s.student
                    });
                }
            });
            
            this.touchDragData = {
                type: 'multipleSeats',
                sourceSeats: Array.from(this.selectedSeats).map(sid => ({ seatId: sid })),
                students: selectedStudents
            };
        } else {
            // 单座位拖拽
            this.touchDragData = {
                type: 'singleSeat',
                studentUuid: seat.student.uuid,
                source: 'seat',
                sourceSeatId: seatId
            };
        }
        
        // 不立即启动拖拽，等待 touchmove 确认
        e.preventDefault();
    }
    
    // 触摸拖拽处理 - 移动
    handleSeatTouchMove(e) {
        if (!this.touchDragData) return;
        
        const touch = e.touches[0];
        const deltaX = Math.abs(touch.clientX - this.touchStartX);
        const deltaY = Math.abs(touch.clientY - this.touchStartY);
        const deltaTime = Date.now() - this.touchStartTime;
        
        // 如果移动距离小于阈值，不启动拖拽（可能是点击）
        if (!this.isTouchDragging && (deltaX < 10 && deltaY < 10)) {
            return;
        }
        
        // 首次移动超过阈值，启动拖拽
        if (!this.isTouchDragging) {
            this.isTouchDragging = true;
            
            // 禁用滚动（触摸设备）
            const classroomGrid = document.getElementById('classroomGrid');
            if (classroomGrid) {
                classroomGrid.classList.add('is-dragging');
            }
            
            // 创建拖拽可视化元素
            this.touchDragElement = document.createElement('div');
            this.touchDragElement.className = 'touch-drag-indicator';
            
            if (this.touchDragData.type === 'multipleSeats') {
                const count = this.touchDragData.students.length;
                this.touchDragElement.innerHTML = `
                    <div style="
                        background: rgba(59, 130, 246, 0.95);
                        color: white;
                        padding: 12px 16px;
                        border-radius: 10px;
                        box-shadow: 0 6px 20px rgba(0, 0, 0, 0.3);
                        font-size: 16px;
                        font-weight: 600;
                        display: flex;
                        align-items: center;
                        gap: 8px;
                        white-space: nowrap;
                        border: 2px solid rgba(255, 255, 255, 0.3);
                        pointer-events: none;
                    ">
                        <span style="font-size: 20px;">📦</span>
                        <span>拖拽 ${count} 个座位</span>
                    </div>
                `;
            } else {
                const seat = this.seats.find(s => s.id === this.touchDragData.sourceSeatId);
                const studentName = seat && seat.student ? seat.student.name : '学生';
                this.touchDragElement.innerHTML = `
                    <div style="
                        background: rgba(37, 99, 235, 0.95);
                        color: white;
                        padding: 12px 16px;
                        border-radius: 10px;
                        box-shadow: 0 6px 20px rgba(0, 0, 0, 0.3);
                        font-size: 16px;
                        font-weight: 600;
                        white-space: nowrap;
                        border: 2px solid rgba(255, 255, 255, 0.3);
                        pointer-events: none;
                    ">
                        ${studentName}
                    </div>
                `;
            }
            
            this.touchDragElement.style.cssText += `
                position: fixed;
                z-index: 10000;
                pointer-events: none;
            `;
            
            document.body.appendChild(this.touchDragElement);
            
            // 标记源座位
            if (this.touchDragData.type === 'multipleSeats') {
                this.selectedSeats.forEach(seatId => {
                    const elem = document.querySelector(`[data-seat-id="${seatId}"]`);
                    if (elem) elem.classList.add('seat-dragging-multi');
                });
            } else {
                const elem = document.querySelector(`[data-seat-id="${this.touchDragData.sourceSeatId}"]`);
                if (elem) elem.classList.add('seat-dragging');
            }
        }
        
        // 更新拖拽元素位置
        if (this.touchDragElement) {
            this.touchDragElement.style.left = (touch.clientX + 15) + 'px';
            this.touchDragElement.style.top = (touch.clientY - 40) + 'px';
        }
        
        // 高亮目标座位
        const targetElement = document.elementFromPoint(touch.clientX, touch.clientY);
        if (targetElement) {
            const targetSeat = targetElement.closest('.seat');
            
            // 清除之前的高亮
            document.querySelectorAll('.seat.drag-over').forEach(el => {
                el.classList.remove('drag-over');
            });
            this.clearMultiDragPreview();
            
            if (targetSeat && targetSeat.dataset.seatId) {
                const targetSeatId = targetSeat.dataset.seatId;
                targetSeat.classList.add('drag-over');
                
                // 显示多选预览
                if (this.touchDragData.type === 'multipleSeats' && this.selectedSeats.size > 0) {
                    const selectedSeatIds = Array.from(this.selectedSeats);
                    this.showMultiDragPreview(targetSeatId, selectedSeatIds);
                }
            }
        }
        
        e.preventDefault();
    }
    
    // 触摸拖拽处理 - 结束
    handleSeatTouchEnd(e) {
        if (!this.touchDragData) return;
        
        // 清除拖拽状态
        document.querySelectorAll('.seat-dragging, .seat-dragging-multi, .drag-over').forEach(el => {
            el.classList.remove('seat-dragging', 'seat-dragging-multi', 'drag-over');
        });
        this.clearMultiDragPreview();
        
        // 移除拖拽可视化元素
        if (this.touchDragElement) {
            this.touchDragElement.remove();
            this.touchDragElement = null;
        }
        
        // 如果是真正的拖拽（而不是点击）
        if (this.isTouchDragging) {
            // 获取释放位置的元素
            const touch = e.changedTouches[0];
            const targetElement = document.elementFromPoint(touch.clientX, touch.clientY);
            
            if (targetElement) {
                const targetSeat = targetElement.closest('.seat');
                
                if (targetSeat && targetSeat.dataset.seatId) {
                    const targetSeatId = targetSeat.dataset.seatId;
                    
                    // 执行放置操作
                    if (this.touchDragData.type === 'multipleSeats') {
                        this.executeMultiDropWithAllSeats(targetSeatId, this.touchDragData);
                    } else if (this.touchDragData.type === 'singleSeat') {
                        this.assignStudentToSeat(
                            this.touchDragData.studentUuid,
                            targetSeatId,
                            this.touchDragData.sourceSeatId
                        );
                    }
                }
            }
        }
        
        // 重置状态
        this.touchDragData = null;
        this.isTouchDragging = false;
        this.touchStartTime = 0;
        
        // 恢复滚动（触摸设备）
        const classroomGrid = document.getElementById('classroomGrid');
        if (classroomGrid) {
            classroomGrid.classList.remove('is-dragging');
        }
        
        e.preventDefault();
    }

    // 创建多选拖拽图像
    createMultiDragImage(e, studentCount) {
        const dragImage = document.createElement('div');
        dragImage.innerHTML = `
            <div style="
                background: rgba(59, 130, 246, 0.95);
                color: white;
                padding: 8px 12px;
                border-radius: 8px;
                box-shadow: 0 4px 12px rgba(0, 0, 0, 0.3);
                font-size: 14px;
                font-weight: 600;
                display: flex;
                align-items: center;
                gap: 6px;
                white-space: nowrap;
                border: 2px solid rgba(255, 255, 255, 0.3);
            ">
                <span style="font-size: 16px;">📦</span>
                <span>拖拽 ${studentCount} 个座位</span>
            </div>
        `;
        
        dragImage.style.cssText = `
            position: absolute;
            top: -1000px;
            left: -1000px;
            z-index: 10000;
            pointer-events: none;
        `;
        
        document.body.appendChild(dragImage);
        
        // 设置自定义拖拽图像
        try {
            e.dataTransfer.setDragImage(dragImage, 60, 20);
        } catch (error) {
            console.warn('设置拖拽图像失败:', error);
        }
        
        // 延时删除临时元素
        requestAnimationFrame(() => {
            setTimeout(() => {
                if (document.body.contains(dragImage)) {
                    document.body.removeChild(dragImage);
                }
            }, 100);
        });
    }

    // 多选拖拽功能
    startMultiDrag(e, draggedSeatId) {
        if (!this.selectedSeats.has(draggedSeatId)) {
            return false; // 不是多选拖拽
        }

        // 收集选中座位的学生信息
        const selectedStudents = [];
        this.selectedSeats.forEach(seatId => {
            const seat = this.seats.find(s => s.id === seatId);
            if (seat && seat.student) {
                selectedStudents.push({
                    seatId: seatId,
                    student: seat.student
                });
            }
        });

        if (selectedStudents.length === 0) {
            return false; // 没有学生可拖拽
        }

        // 设置拖拽数据
        const dragData = {
            type: 'multipleSeats',
            students: selectedStudents,
            sourceSeats: Array.from(this.selectedSeats)
        };

        e.dataTransfer.setData('text/plain', JSON.stringify(dragData));
        
        // 创建自定义拖拽图像 - 小手图标
        const dragImage = document.createElement('div');
        dragImage.innerHTML = '✋';
        dragImage.style.cssText = `
            position: absolute;
            top: -1000px;
            left: -1000px;
            width: 40px;
            height: 40px;
            font-size: 24px;
            display: flex;
            align-items: center;
            justify-content: center;
            background: rgba(255, 255, 255, 0.9);
            border: 2px solid #3b82f6;
            border-radius: 8px;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.3);
            z-index: 10000;
        `;
        document.body.appendChild(dragImage);
        
        // 设置自定义拖拽图像
        try {
            e.dataTransfer.setDragImage(dragImage, 20, 20);
        } catch (error) {
            console.warn('设置拖拽图像失败:', error);
        }
        
        // 延时删除临时元素
        requestAnimationFrame(() => {
            setTimeout(() => {
                if (document.body.contains(dragImage)) {
                    document.body.removeChild(dragImage);
                }
            }, 100);
        });
        
        // 添加视觉反馈
        this.selectedSeats.forEach(seatId => {
            const seatElement = document.querySelector(`[data-seat-id="${seatId}"]`);
            if (seatElement) {
                seatElement.classList.add('seat-dragging-multi');
            }
        });

        return true;
    }

    // 处理多选拖拽结束
    endMultiDrag() {
        document.querySelectorAll('.seat-dragging-multi').forEach(element => {
            element.classList.remove('seat-dragging-multi');
        });
        document.querySelectorAll('.seat-drag-preview').forEach(element => {
            element.classList.remove('seat-drag-preview');
        });
    }

    // 显示多选拖拽预览
    showMultiDragPreview(targetSeatId, selectedSeatIds) {
        // 清除之前的预览
        this.clearMultiDragPreview();
        
        if (!targetSeatId || selectedSeatIds.length === 0) return;
        
        // 检查目标位置是否有效
        const dropResult = this.checkMultiDropTarget(targetSeatId, selectedSeatIds);
        if (!dropResult) return;
        
        const { positions: targetPositions } = dropResult;
        
        // 只在目标位置显示预览
        targetPositions.forEach(positionMapping => {
            const targetSeatElement = document.querySelector(`[data-seat-id="${positionMapping.targetSeatId}"]`);
            
            if (targetSeatElement) {
                // 获取原始座位的学生信息
                const originalSeat = this.seats.find(s => s.id === positionMapping.originalSeatId);
                
                if (originalSeat && originalSeat.student) {
                    // 添加预览样式到目标座位
                    targetSeatElement.classList.add('seat-drag-preview');
                    
                    // 隐藏目标座位原有的学生姓名（如果有）
                    const existingName = targetSeatElement.querySelector('.student-name-display');
                    if (existingName) {
                        existingName.style.opacity = '0.3';
                    }
                    
                    // 创建预览学生姓名显示（在目标位置）
                    const previewElement = document.createElement('div');
                    previewElement.className = 'drag-preview-student';
                    previewElement.textContent = '→ ' + originalSeat.student.name;
                    previewElement.style.cssText = `
                        position: absolute;
                        top: 50%;
                        left: 50%;
                        transform: translate(-50%, -50%);
                        font-size: 1.1rem;
                        font-weight: bold;
                        color: #1d4ed8;
                        background: rgba(255, 255, 255, 0.98);
                        padding: 5px 10px;
                        border-radius: 5px;
                        border: 2px solid #3b82f6;
                        box-shadow: 0 3px 10px rgba(59, 130, 246, 0.5);
                        pointer-events: none;
                        z-index: 2000;
                        white-space: nowrap;
                        max-width: 90%;
                        overflow: hidden;
                        text-overflow: ellipsis;
                    `;
                    
                    targetSeatElement.appendChild(previewElement);
                }
            }
        });
    }

    // 清除多选拖拽预览
    clearMultiDragPreview() {
        // 移除所有预览样式
        document.querySelectorAll('.seat-drag-preview').forEach(element => {
            element.classList.remove('seat-drag-preview');
            
            // 恢复原有学生姓名的显示
            const existingName = element.querySelector('.student-name-display');
            if (existingName) {
                existingName.style.opacity = '';
            }
        });
        
        // 移除所有预览学生姓名
        document.querySelectorAll('.drag-preview-student').forEach(element => {
            element.remove();
        });
    }

    // 计算选中座位的相对位置布局
    calculateRelativeLayout(selectedSeatIds) {
        if (selectedSeatIds.length === 0) return { layout: [], bounds: {} };
        
        // 获取所有选中座位的坐标
        const positions = selectedSeatIds.map(seatId => {
            const seat = this.seats.find(s => s.id === seatId);
            return seat ? { seatId, row: seat.row, col: seat.col } : null;
        }).filter(pos => pos !== null);
        
        if (positions.length === 0) return { layout: [], bounds: {} };
        
        // 找到边界
        const minRow = Math.min(...positions.map(p => p.row));
        const maxRow = Math.max(...positions.map(p => p.row));
        const minCol = Math.min(...positions.map(p => p.col));
        const maxCol = Math.max(...positions.map(p => p.col));
        
        // 计算相对位置（以最小坐标为原点）
        const layout = positions.map(pos => ({
            seatId: pos.seatId,
            relativeRow: pos.row - minRow,
            relativeCol: pos.col - minCol
        }));
        
        return {
            layout,
            bounds: {
                width: maxCol - minCol + 1,
                height: maxRow - minRow + 1,
                minRow,
                maxRow,
                minCol,
                maxCol
            }
        };
    }

    // 智能调整目标位置以适应边界
    adjustTargetPosition(targetRow, targetCol, bounds) {
        let adjustedRow = targetRow;
        let adjustedCol = targetCol;
        
        // 检查并调整行位置
        const maxPossibleRow = targetRow + bounds.height - 1;
        if (maxPossibleRow >= this.rows) {
            adjustedRow = this.rows - bounds.height;
        }
        if (adjustedRow < 0) {
            adjustedRow = 0;
        }
        
        // 检查并调整列位置
        const maxPossibleCol = targetCol + bounds.width - 1;
        if (maxPossibleCol >= this.cols) {
            adjustedCol = this.cols - bounds.width;
        }
        if (adjustedCol < 0) {
            adjustedCol = 0;
        }
        
        return { adjustedRow, adjustedCol };
    }

    // 检查多选拖拽的目标区域（保持相对位置，智能边界调整）
    checkMultiDropTarget(targetSeatId, selectedSeatIds) {
        const targetSeat = this.seats.find(s => s.id === targetSeatId);
        if (!targetSeat) return false;

        const targetRow = targetSeat.row;
        const targetCol = targetSeat.col;

        // 计算选中座位的相对布局
        const relativeLayout = this.calculateRelativeLayout(selectedSeatIds);
        if (relativeLayout.layout.length === 0) return false;

        const { layout, bounds } = relativeLayout;
        
        // 检查布局是否能够适应教室尺寸
        if (bounds.width > this.cols || bounds.height > this.rows) {
            return false; // 选中区域本身就超过了教室尺寸
        }

        // 智能调整目标位置
        const { adjustedRow, adjustedCol } = this.adjustTargetPosition(
            targetRow, targetCol, bounds
        );

        // 计算每个座位的新位置和检查冲突
        const targetPositions = [];
        const displacedStudents = [];
        
        for (const item of layout) {
            const newRow = adjustedRow + item.relativeRow;
            const newCol = adjustedCol + item.relativeCol;
            const newSeatId = `${newRow}-${newCol}`;
            
            const newSeat = this.seats.find(s => s.id === newSeatId);
            if (!newSeat || newSeat.isDeleted) {
                return false; // 座位不存在或已删除
            }

            targetPositions.push({
                originalSeatId: item.seatId,
                targetSeatId: newSeatId,
                row: newRow,
                col: newCol
            });

            // 如果目标座位上有学生且不在选中列表中，记录被替换的学生
            if (newSeat.student && !selectedSeatIds.includes(newSeatId)) {
                displacedStudents.push({
                    student: newSeat.student,
                    targetSeatId: newSeatId,
                    originalSeatId: item.seatId  // 添加原座位ID，用于位置交换
                });
            }
        }

        return {
            positions: targetPositions,
            displacedStudents: displacedStudents,
            layout: relativeLayout,
            wasAdjusted: adjustedRow !== targetRow || adjustedCol !== targetCol,
            adjustedPosition: { row: adjustedRow, col: adjustedCol }
        };
    }

    // 执行多选拖拽放置（新版本，支持空座位）
    executeMultiDropWithAllSeats(targetSeatId, dragData) {
        // 从dragData中提取座位ID列表
        let selectedSeatIds = [];
        if (dragData.sourceSeats) {
            // 新格式：处理对象数组，提取seatId
            if (Array.isArray(dragData.sourceSeats) && dragData.sourceSeats.length > 0) {
                if (typeof dragData.sourceSeats[0] === 'object' && dragData.sourceSeats[0].seatId) {
                    selectedSeatIds = dragData.sourceSeats.map(seat => seat.seatId);
                } else {
                    // 旧格式：直接使用座位ID数组
                    selectedSeatIds = dragData.sourceSeats;
                }
            }
        }
        
        if (selectedSeatIds.length === 0) return false;
        
        const dropResult = this.checkMultiDropTarget(targetSeatId, selectedSeatIds);
        
        if (!dropResult) {
            // 分析具体原因并提供更好的错误信息
            const relativeLayout = this.calculateRelativeLayout(selectedSeatIds);
            if (relativeLayout.layout.length === 0) {
                alert('没有选中有效的座位');
            } else {
                const { bounds } = relativeLayout;
                if (bounds.width > this.cols || bounds.height > this.rows) {
                    alert(`选中区域过大无法放置：\n选中区域: ${bounds.width}列 × ${bounds.height}行\n教室尺寸: ${this.cols}列 × ${this.rows}行\n\n请选择较小的区域或调整教室布局`);
                } else {
                    alert('目标位置无法容纳选中的座位，请选择其他位置');
                }
            }
            return false;
        }

        const { positions: targetPositions, displacedStudents, wasAdjusted, adjustedPosition } = dropResult;

        this.addToHistory('seatArrangement', { seats: this.seats });

        // 创建座位ID到学生的映射（包括空座位）
        const seatMap = new Map();
        selectedSeatIds.forEach(seatId => {
            const seat = this.seats.find(s => s.id === seatId);
            if (seat) {
                seatMap.set(seatId, seat.student || null); // null表示空座位
            }
        });

        // 清空原座位
        selectedSeatIds.forEach(seatId => {
            const seat = this.seats.find(s => s.id === seatId);
            if (seat) {
                seat.student = null;
            }
        });

        // 处理被顶替的学生
        const displacedStudentMap = new Map();
        displacedStudents.forEach(displaced => {
            displacedStudentMap.set(displaced.targetSeatId, displaced.student);
        });

        // 在新位置放置学生
        targetPositions.forEach(positionMapping => {
            const targetSeat = this.seats.find(s => s.id === positionMapping.targetSeatId);
            const originalStudent = seatMap.get(positionMapping.originalSeatId);
            
            if (targetSeat) {
                targetSeat.student = originalStudent; // 可能是null（空座位）
            }
        });

        // 将被顶替的学生放到原来的位置（智能分配）
        const availableOriginalSeats = new Set(selectedSeatIds.map(id => this.seats.find(s => s.id === id)).filter(seat => seat));
        
        displacedStudents.forEach(displaced => {
            const preferredSeat = this.seats.find(s => s.id === displaced.originalSeatId);
            
            // 如果首选座位可用且为空，直接分配
            if (preferredSeat && !preferredSeat.student && availableOriginalSeats.has(preferredSeat)) {
                preferredSeat.student = displaced.student;
                availableOriginalSeats.delete(preferredSeat);
            } else {
                // 否则找任何可用的原座位
                const availableSeat = Array.from(availableOriginalSeats).find(seat => !seat.student);
                if (availableSeat) {
                    availableSeat.student = displaced.student;
                    availableOriginalSeats.delete(availableSeat);
                } else {
                    // 如果所有原座位都被占用，寻找任何空座位
                    const emptySeat = this.seats.find(seat => !seat.student);
                    if (emptySeat) {
                        emptySeat.student = displaced.student;
                    }
                    // 如果没有空座位，学生将保持在未安排状态（自动回到学生列表）
                }
            }
        });

        // 清除多选状态
        this.clearSelection();

        // 更新界面（使用增量更新提升性能）
        this.saveData();
        this.renderClassroom(false);
        this.renderStudentList();
        this.updateStats();
        this.applyCurrentFilter();

        // 提供用户反馈
        let feedbackMessage = '';
        if (wasAdjusted) {
            feedbackMessage += `位置已自动调整到边界内。`;
        }
        if (displacedStudents.length > 0) {
            feedbackMessage += `移动完成：${displacedStudents.length} 个学生发生了位置对调`;
        } else {
            feedbackMessage += `多选移动完成：移动了 ${selectedSeatIds.length} 个座位`;
        }
        
        if (feedbackMessage) {
            console.log(feedbackMessage);
            // 可选：显示用户提示
            // alert(feedbackMessage);
        }

        return true;
    }

    // 执行多选拖拽放置（旧版本，保持兼容性）
    executeMultiDrop(targetSeatId, studentsData) {
        // 提取选中座位的ID列表
        const selectedSeatIds = studentsData.map(data => data.seatId);
        const dropResult = this.checkMultiDropTarget(targetSeatId, selectedSeatIds);
        
        if (!dropResult) {
            // 分析具体原因并提供更好的错误信息
            const relativeLayout = this.calculateRelativeLayout(selectedSeatIds);
            if (relativeLayout.layout.length === 0) {
                alert('没有选中有效的座位');
            } else {
                const { bounds } = relativeLayout;
                if (bounds.width > this.cols || bounds.height > this.rows) {
                    alert(`选中区域过大无法放置：\n选中区域: ${bounds.width}列 × ${bounds.height}行\n教室尺寸: ${this.cols}列 × ${this.rows}行\n\n请选择较小的区域或调整教室布局`);
                } else {
                    alert('目标位置无法放置选中的座位，请尝试其他位置');
                }
            }
            return false;
        }

        const { positions: targetPositions, displacedStudents, wasAdjusted, adjustedPosition } = dropResult;

        this.addToHistory('seatArrangement', { seats: this.seats });

        // 创建原始座位ID到学生的映射
        const studentMap = new Map();
        studentsData.forEach(data => {
            studentMap.set(data.seatId, data.student);
        });

        // 先清空原座位
        studentsData.forEach(data => {
            const seat = this.seats.find(s => s.id === data.seatId);
            if (seat) {
                seat.student = null;
            }
        });

        // 按照相对位置关系放置到新位置
        targetPositions.forEach(positionMapping => {
            const originalStudent = studentMap.get(positionMapping.originalSeatId);
            if (originalStudent) {
                const newSeat = this.seats.find(s => s.id === positionMapping.targetSeatId);
                if (newSeat) {
                    newSeat.student = originalStudent;
                }
            }
        });

        // 处理被替换的学生 - 将他们移动到原来被移动学生的位置
        if (displacedStudents.length > 0) {
            // 获取原始座位ID列表
            const originalSeatIds = studentsData.map(data => data.seatId);
            
            // 为被替换的学生分配到原始位置
            displacedStudents.forEach((displaced, index) => {
                if (index < originalSeatIds.length) {
                    const originalSeat = this.seats.find(s => s.id === originalSeatIds[index]);
                    if (originalSeat) {
                        originalSeat.student = displaced.student;
                    }
                } else {
                    // 如果被替换的学生数量超过原始位置数量，寻找空座位
                    const emptySeat = this.seats.find(seat => !seat.student);
                    if (emptySeat) {
                        emptySeat.student = displaced.student;
                    }
                }
            });
        }

        this.clearSelection();
        this.saveData();
        this.renderClassroom(false); // 使用增量更新
        this.renderStudentList();
        this.updateStats();
        this.applyCurrentFilter();

        // 向用户提供操作反馈
        let feedbackMessage = '';
        
        if (wasAdjusted) {
            feedbackMessage += `位置已自动调整到合适区域。`;
        }
        
        if (displacedStudents.length > 0) {
            feedbackMessage += `位置对调完成：移动了 ${studentsData.length} 个学生，对调了 ${displacedStudents.length} 个学生`;
        } else {
            feedbackMessage += `多选移动完成：移动了 ${studentsData.length} 个学生到空座位`;
        }
        
        if (feedbackMessage) {
            console.log(feedbackMessage);
        }

        return true;
    }

    // 座位轮换功能
    rotateSeats(direction) {
        // 检查是否有已安排的学生
        const occupiedSeats = this.seats.filter(seat => seat.student);
        if (occupiedSeats.length === 0) {
            alert('当前没有已安排座位的学生，无法进行轮换');
            return;
        }

        // 添加到历史记录
        this.addToHistory('seatArrangement', { seats: this.seats });

        // 执行轮换
        switch (direction) {
            case 'rowLeft':
                this.rotateRowsLeft();
                break;
            case 'rowRight':
                this.rotateRowsRight();
                break;
            case 'colForward':
                this.rotateColumnsForward();
                break;
            case 'colBackward':
                this.rotateColumnsBackward();
                break;
        }

        // 更新界面（使用增量更新提升性能）
        this.saveData();
        this.renderClassroom(false);
        this.renderStudentList();
        this.updateStats();
        this.applyCurrentFilter();
        
        // 记录操作日志
        const directionText = {
            'rowLeft': '整排向左轮换',
            'rowRight': '整排向右轮换', 
            'colForward': '整列向前轮换',
            'colBackward': '整列向后轮换'
        };
        console.log(`执行座位轮换: ${directionText[direction]}`);
    }

    // 按排向左轮换：每一排的学生都向左移动一位，最左边的移到最右边
    rotateRowsLeft() {
        for (let row = 0; row < this.rows; row++) {
            // 获取当前行的所有座位，排除已删除的座位
            const rowSeats = this.seats.filter(seat => seat.row === row && !seat.isDeleted);
            rowSeats.sort((a, b) => a.col - b.col); // 按列排序
            
            // 提取学生信息
            const students = rowSeats.map(seat => seat.student);
            
            // 向左轮换：第一个学生移到最后，其他学生向前移动
            if (students.some(student => student !== null)) {
                const rotatedStudents = [...students.slice(1), students[0]];
                
                // 重新分配学生到座位
                rowSeats.forEach((seat, index) => {
                    seat.student = rotatedStudents[index];
                });
            }
        }
    }

    // 按排向右轮换：每一排的学生都向右移动一位，最右边的移到最左边
    rotateRowsRight() {
        for (let row = 0; row < this.rows; row++) {
            // 获取当前行的所有座位，排除已删除的座位
            const rowSeats = this.seats.filter(seat => seat.row === row && !seat.isDeleted);
            rowSeats.sort((a, b) => a.col - b.col); // 按列排序
            
            // 提取学生信息
            const students = rowSeats.map(seat => seat.student);
            
            // 向右轮换：最后一个学生移到最前，其他学生向后移动
            if (students.some(student => student !== null)) {
                const rotatedStudents = [students[students.length - 1], ...students.slice(0, -1)];
                
                // 重新分配学生到座位
                rowSeats.forEach((seat, index) => {
                    seat.student = rotatedStudents[index];
                });
            }
        }
    }

    // 按列向前轮换：每一列的学生都向前移动一位，最前面的移到最后面
    rotateColumnsForward() {
        for (let col = 0; col < this.cols; col++) {
            // 获取当前列的所有座位，排除已删除的座位
            const colSeats = this.seats.filter(seat => seat.col === col && !seat.isDeleted);
            colSeats.sort((a, b) => a.row - b.row); // 按行排序（row=0是前排）
            
            // 提取学生信息
            const students = colSeats.map(seat => seat.student);
            
            // 向前轮换：第一个学生（前排）移到最后，其他学生向前移动
            if (students.some(student => student !== null)) {
                const rotatedStudents = [...students.slice(1), students[0]];
                
                // 重新分配学生到座位
                colSeats.forEach((seat, index) => {
                    seat.student = rotatedStudents[index];
                });
            }
        }
    }

    // 按列向后轮换：每一列的学生都向后移动一位，最后面的移到最前面
    rotateColumnsBackward() {
        for (let col = 0; col < this.cols; col++) {
            // 获取当前列的所有座位，排除已删除的座位
            const colSeats = this.seats.filter(seat => seat.col === col && !seat.isDeleted);
            colSeats.sort((a, b) => a.row - b.row); // 按行排序（row=0是前排）
            
            // 提取学生信息
            const students = colSeats.map(seat => seat.student);
            
            // 向后轮换：最后一个学生（后排）移到最前，其他学生向后移动
            if (students.some(student => student !== null)) {
                const rotatedStudents = [students[students.length - 1], ...students.slice(0, -1)];
                
                // 重新分配学生到座位
                colSeats.forEach((seat, index) => {
                    seat.student = rotatedStudents[index];
                });
            }
        }
    }

}

const app = new SeatingApp();