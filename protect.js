/**
 * 排座宝 - 页面保护层
 * Copyright © 2025 排座宝 - 版权所有
 * 
 * 警告：此代码受版权保护，未经授权不得复制、修改或分发
 * Warning: This code is protected by copyright law. Unauthorized copying, modification or distribution is prohibited.
 */

// 增强的右键和快捷键保护
(function() {
    'use strict';
    
    // 全面的右键拦截
    const blockRightClick = function(e) {
        e = e || window.event;
        
        // 检测所有右键相关事件
        if (e.button === 2 || e.which === 3 || e.type === 'contextmenu' ||
            (e.type === 'mousedown' && e.button === 2) ||
            (e.type === 'mouseup' && e.button === 2) ||
            (e.buttons && (e.buttons & 2))) {
            
            // 阻止所有默认行为
            e.preventDefault();
            e.stopPropagation();
            e.stopImmediatePropagation();
            
            if (e.returnValue !== undefined) e.returnValue = false;
            if (e.cancelBubble !== undefined) e.cancelBubble = true;
            
            return false;
        }
    };
    
    // 增强的快捷键拦截
    const blockKeys = function(e) {
        const keyCode = e.keyCode || e.which;
        const key = e.key ? e.key.toLowerCase() : '';
        const ctrl = e.ctrlKey;
        const shift = e.shiftKey;
        const meta = e.metaKey;
        
        // 定义被禁用的快捷键组合
        const blockedCombinations = [
            // 开发者工具
            keyCode === 123, // F12
            key === 'f12',
            (ctrl && shift && keyCode === 73), // Ctrl+Shift+I
            (ctrl && shift && key === 'i'),
            (ctrl && shift && keyCode === 74), // Ctrl+Shift+J
            (ctrl && shift && key === 'j'),
            (ctrl && shift && keyCode === 67), // Ctrl+Shift+C
            (ctrl && shift && key === 'c'),
            (ctrl && keyCode === 85), // Ctrl+U (查看源码)
            (ctrl && key === 'u'),
            
            // Safari/Mac 快捷键
            (meta && keyCode === 73), // Cmd+Option+I
            (meta && key === 'i'),
            (meta && keyCode === 74), // Cmd+Option+J
            (meta && key === 'j'),
            (meta && keyCode === 85), // Cmd+U
            (meta && key === 'u'),
            
            // Firefox 快捷键
            (ctrl && shift && keyCode === 75), // Ctrl+Shift+K
            (ctrl && shift && key === 'k')
        ];
        
        // 检查是否命中禁用列表
        if (blockedCombinations.some(condition => condition)) {
            e.preventDefault();
            e.stopPropagation();
            e.stopImmediatePropagation();
            
            return false;
        }
    };
    
    // 检查是否为输入字段
    function isInputField(element) {
        if (!element) return false;
        const inputTags = ['INPUT', 'TEXTAREA', 'SELECT'];
        return inputTags.includes(element.tagName) || 
               element.contentEditable === 'true';
    }
    
    // 多层次事件绑定
    function bindProtection() {
        // 右键保护
        const rightClickEvents = ['contextmenu', 'mousedown', 'mouseup', 'auxclick'];
        rightClickEvents.forEach(eventType => {
            document.addEventListener(eventType, blockRightClick, true);
        });
        
        // 键盘保护
        const keyboardEvents = ['keydown', 'keypress'];
        keyboardEvents.forEach(eventType => {
            document.addEventListener(eventType, blockKeys, true);
        });
        
        // 直接属性绑定
        document.oncontextmenu = blockRightClick;
        document.onkeydown = blockKeys;
    }
    
    // 立即绑定保护
    bindProtection();
    
    // 页面加载后重新绑定
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', bindProtection);
    }
    window.addEventListener('load', bindProtection);
    
    // 禁用拖拽（除了应用内的拖拽功能）
    document.addEventListener('dragstart', function(e) {
        // 允许应用内的学生拖拽功能
        if (e.target.classList.contains('student-item') || 
            e.target.closest('.student-item') ||
            e.target.classList.contains('student-name-display') ||
            e.target.closest('.student-name-display') ||
            e.target.classList.contains('seat') ||
            e.target.closest('.seat')) {
            return true;
        }
        e.preventDefault();
        return false;
    });
    
})();

// 版权保护措施
(function() {
    'use strict';
    
    // 控制台版权警告
    function showCopyrightWarning() {
        const styles = [
            'color: #e74c3c',
            'font-size: 20px',
            'font-weight: bold',
            'text-shadow: 2px 2px 4px rgba(0,0,0,0.3)'
        ].join(';');
        
        const styles2 = [
            'color: #c0392b',
            'font-size: 16px',
            'font-weight: bold'
        ].join(';');
        
        const styles3 = [
            'color: #7f8c8d',
            'font-size: 14px'
        ].join(';');
        
        console.clear();
        console.log('%c⚠️ 排座宝 - 版权保护 ⚠️', styles);
        console.log('%c本软件受版权保护，未经授权不得复制、修改或分发', styles2);
        console.log('%c违反版权法将承担法律责任', styles2);
        console.log('%c© ' + new Date().getFullYear() + ' 排座宝 - 版权所有', styles3);
        console.log('%c如需使用请联系：小红书 @排座宝', styles3);
    }
    
    // 防止通过 iframe 嵌入
    function preventFraming() {
        if (window.top !== window.self) {
            try {
                window.top.location = window.self.location;
            } catch (e) {
                // 如果跨域限制，则隐藏页面
                document.body.style.display = 'none';
            }
        }
    }
    
    // 检测自动化工具
    function detectAutomation() {
        // 检测 Selenium WebDriver
        if (window.navigator.webdriver || 
            window.document.documentElement.getAttribute('webdriver') ||
            window.callPhantom || window._phantom) {
            showCopyrightWarning();
            return true;
        }
        return false;
    }
    
    // 开发者工具检测
    function detectDevTools() {
        const threshold = 160;
        let devtoolsOpen = false;
        
        setInterval(() => {
            const heightDiff = window.outerHeight - window.innerHeight;
            const widthDiff = window.outerWidth - window.innerWidth;
            
            if (heightDiff > threshold || widthDiff > threshold) {
                if (!devtoolsOpen) {
                    devtoolsOpen = true;
                    showCopyrightWarning();
                }
            } else {
                devtoolsOpen = false;
            }
        }, 1000);
    }
    
    // 初始化保护
    function initProtection() {
        showCopyrightWarning();
        preventFraming();
        detectAutomation();
        detectDevTools();
    }
    
    // 页面加载时启动保护
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', initProtection);
    } else {
        initProtection();
    }
    
    // 页面可见性变化时重新检查
    document.addEventListener('visibilitychange', () => {
        if (!document.hidden) {
            setTimeout(showCopyrightWarning, 500);
        }
    });
    
})();

// 水印保护
(function() {
    'use strict';
    
    // 创建版权水印
    function createWatermark() {
        // 检查是否已存在水印
        if (document.getElementById('copyright-watermark')) {
            return;
        }
        
        const watermark = document.createElement('div');
        watermark.id = 'copyright-watermark';
        watermark.style.cssText = `
            position: fixed;
            bottom: 10px;
            right: 10px;
            font-size: 12px;
            color: rgba(0, 0, 0, 0.15);
            pointer-events: none;
            user-select: none;
            z-index: 999999;
            font-family: Arial, sans-serif;
        `;
        watermark.textContent = '© 2025 排座宝';
        
        document.body.appendChild(watermark);
        
        // 监控水印是否被删除
        const observer = new MutationObserver((mutations) => {
            mutations.forEach((mutation) => {
                mutation.removedNodes.forEach((node) => {
                    if (node.id === 'copyright-watermark') {
                        setTimeout(createWatermark, 100);
                    }
                });
            });
        });
        
        observer.observe(document.body, {
            childList: true,
            subtree: true
        });
    }
    
    // 页面加载后添加水印
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', createWatermark);
    } else {
        createWatermark();
    }
    
})();
