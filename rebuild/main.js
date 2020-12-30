"use strict"

var app = new Vue({
    el: '#app',
    data: {
        names: [],
        isDragIn: false,
        startTime: null,
        logs: {},
        step: 0,
        totalStep: 7,
        showMultipleFileWarning: false,
        selected: ['styles', 'workbook'],
        options: [
            { text: '清除自定义样式', value: 'styles' },
            { text: '清除命名区域', value: 'workbook' },
            { text: 'PNG→JPG 图片批量转换（未完成）', value: 'png' }
        ]
    },
    methods: {
        onDragenter: function () {
            this.isDragIn = true;
        },
        onDragleave: function () {
            this.isDragIn = false;
        },
        onGetFiles: function (e) {
            // 判断是选择还是拖入的
            if (e.type == 'drop') {
                var files = e.dataTransfer.files;
                this.isDragIn = false;
            } else {
                var files = e.target.files;
            }
            // 初始化
            this.startTime = new Date().getTime();
            this.logs = {};
            this.alerts = {};
            this.step = 0;
            this.totalStep = (files.length) * 7;
            if (files.length > 1) {
                this.showMultipleFileWarning = true;
            } else {
                this.showMultipleFileWarning = false;
            }
            // 所有文件并行处理
            for (var i = 0; i < files.length; i++) {
                handleFile(files[i]);
            }
        },
        log: function (key, value, variant) {
            console.log(value)
            if (!this.logs[key]) {
                this.logs[key] = { step: 0 }
            }
            var timePast = '耗时 ' + (((new Date).getTime() - this.startTime) / 1000).toFixed(1) + 's';
            this.logs[key].timePast = timePast;
            this.logs[key].value = value;
            this.logs[key].variant = variant;
            // 由于vue限制，对于 logs 深层数值的多次更新，视图只显示初始化之后的第一次更新，
            // 因此通过更新位于数据顶层的进度条，即 data 里的 step，使深层日志的每次更新也都能被立即显示出来。
            this.step += 1
        }

    }
})

// 处理文件
async function handleFile(f) {
    var options = {}
    app.selected.forEach(element => {
        options[element] = true;
    });
    console.log(options)
    app.log(f.name, '读取文件。');
    var zip = await JSZip.loadAsync(f).catch(function () {
        var msg = '解压错误，请确认文件是 xlsx 格式，如果是旧版本的 xls 文件，请先在 Excel 中另存为 xlsx 格式。';
        app.log(f.name, msg, 'text-danger')
        app.step += 5;
    });

    if (!zip) { return }

    if (options.styles) {
        if (!(zip.file('xl/styles.xml'))) {
            var msg = '文件中未找到 /xl/styles.xml。'
            app.log(f.name, msg, 'text-danger');
            app.step += 5;
            return
        }
        // 读取 styles.xml 文件
        app.log(f.name, '读取 styles.xml 文件。');
        var stylesXML = await zip.file('xl/styles.xml').async('string');
        // 清除自定义样式
        app.log(f.name, '清除自定义样式。')
        stylesXML = stylesXML.replace(/<numFmts.*<\/numFmts>/, '');
        stylesXML = stylesXML.replace(/<cellStyleXfs.*<\/cellStyleXfs>/, '');
        stylesXML = stylesXML.replace(/<cellStyles.*<\/cellStyles>/, '');
        // 修改后的文件写入zip
        zip.file('xl/styles.xml', stylesXML);
    } else {
        app.step += 2;
    }

    if (options.workbook) {
        if (!(zip.file('xl/workbook.xml'))) {
            var msg = '文件中未找到 /xl/workbook.xml。'
            app.log(f.name, msg, 'text-danger');
            app.step += 3;
            return
        }
        // 读取 workbook.xml 文件
        app.log(f.name, '读取 workbook.xml 文件。');
        var workbookXML = await zip.file('xl/workbook.xml').async('string');
        // 清除自定义样式
        app.log(f.name, '清除命名区域。')
        workbookXML = workbookXML.replace(/<definedNames.*<\/definedNames>/, '');
        // 修改后的文件写入zip
        zip.file('xl/workbook.xml', workbookXML);
    } else {
        app.step += 2;
    }

    // 重新打包
    app.log(f.name, '重新压缩文件。');
    zip.generateAsync({
        type: 'blob',
        compression: 'DEFLATE',
        compressionOptions: { level: 6 }
    }).then(function (blob) {
        // 输出到FileSaver.min.js
        saveAs(blob, f.name);
        var msg = (f.size / 1024).toFixed(1) + 'KB → ' + (blob.size / 1024).toFixed(1) + 'KB，压缩比例：' + (100 - blob.size / f.size * 100).toFixed(1) + '%。'
        app.log(f.name, '√ 完成！' + msg, 'text-success');
    });

}

// 获取 00:00:00 格式时间
function getHHMMSS(t) {
    if (!t) {
        t = new Date()
    }
    return double(t.getHours()) + ':' + double(t.getMinutes()) + ':' + double(t.getSeconds()) + ':' + (t.getMilliseconds())

    function double(a) {
        return Number(a) < 10 ? ('0' + String(a)) : a
    }
}