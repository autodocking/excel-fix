"use strict"

var app = new Vue({
    el: '#app',
    data: {
        names: [],
        isDragIn: false,
        startTime: null,
        logs: {},
        step: 0,
        totalStep: 5,
        showMultipleFileWarning: false
    },
    methods: {
        onDragenter: function (e) {
            this.isDragIn = true;
        },
        onDragleave: function (e) {
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
            this.totalStep = (files.length) * 5;
            if (files.length > 1) {
                this.showMultipleFileWarning = true;
            } else {
                this.showMultipleFileWarning = false;
            }
            // 所有文件并行处理
            for (var i = 0; i < files.length; i++) {
                handleFile(files[i], i);
            }
        },
        log: function (key, value, variant) {
            if (!this.logs[key]) {
                this.logs[key] = { step: 0 }
            }
            var timePast = '耗时 ' + (((new Date).getTime() - this.startTime) / 1000).toFixed(1) + 's';
            this.logs[key].timePast = timePast;
            this.logs[key].value = value;
            this.logs[key].variant = variant;
            // 由于vue限制，对 logs 深层数值的多次更新，视图只显示初始化之后的第一次更新
            // 因此通过更新处于数据最外层的进度条 step，使深层日志的每次更新也都能被立即显示
            this.step += 1
        },

    }
})

// 处理文件
function handleFile(f, i) {

    app.log(f.name, '读取文件。')

    JSZip.loadAsync(f)
        .then(function (zip) {

            app.log(f.name, '解压文件。')

            if (!(zip.file('xl/styles.xml'))) {
                var msg = '文件已解压，但是未找到 /xl/styles.xml，这可能不是一个 Excel 文件。'
                // app.log(f.name, msg)
                app.log(f.name, msg, 'text-danger');
                app.step += 2;
                return;
            }
            zip.file('xl/styles.xml').async('string').then(function (temp) {
                // 清除自定义样式
                app.log(f.name, '清除自定义样式。')
                temp = temp.replace(/<numFmts.*<\/numFmts>/, '');
                temp = temp.replace(/<cellStyleXfs.*<\/cellStyleXfs>/, '');
                temp = temp.replace(/<cellStyles.*<\/cellStyles>/, '');
                // 修改后的文件写入zip
                zip.file('xl/styles.xml', temp);

                app.log(f.name, '重新压缩文件。');
                zip.generateAsync({
                    type: 'blob',
                    compression: 'DEFLATE',
                    compressionOptions: { level: 6 }
                }).then(function (blob) {
                    // 输出到FileSaver.min.js
                    saveAs(blob, f.name);
                    app.log(f.name, '√ 完成！', 'text-success');
                });
            });
        }, function (e) {
            var msg = '解压错误，请确认文件是 xlsx 格式，如果是旧版本的 xls 文件，请先在 Excel 中另存为 xlsx 格式。';
            app.log(f.name, msg, 'text-danger')
            app.step += 3;
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