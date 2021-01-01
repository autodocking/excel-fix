"use strict"

var app = new Vue({
    el: '#app',
    data: {
        files: [],
        imageQuality: 80,
        isDragIn: false,
        startTime: 0,
        logs: {},
        step: 0,
        showMultipleFileWarning: false,
        selected: ['styles'],
        options: [
            { text: '清除自定义样式', value: 'styles' },
            { text: '清除自定义名称', value: 'workbook' },
            { text: '压缩图片（还支持批量压缩 pptx、docx、zip 文件中的图片）', value: 'png' },
            { text: '同时处理多个文件时，将结果打包成一个 zip 文件（未完成）', value: 'zip', disabled: true }
        ],
        imagesOptionsSelected: ['png2Jpeg'],
        imagesOptions: [
            { text: '将 png 格式的图片转化为 jpg 格式', value: 'png2Jpeg' },
            { text: '将 png 透明区域转化为白色，而不是黑色（未完成）', value: 'whiteBackground', disabled: true },
            { text: '同时修改扩展名（未完成）', value: 'changeExtensionName', disabled: true },
            { text: '显示出所有图片，由我来决定压缩哪些（未完成）', value: 'showThumbnails', disabled: true },
            { text: '将所有图片转化为 webp 格式（待验证兼容性）', value: 'all2Webp', disabled: true }
        ],
        setting: {
            quality: 0.8,
            size: 1,
        }
    },
    methods: {
        onDragenter() {
            this.isDragIn = true;
        },
        onDragleave() {
            this.isDragIn = false;
        },
        onGetFiles(e) {
            if (!this.selectAtLeastOne) {
                this.isDragIn = false;
                return
            }
            // 判断是选择还是拖入的
            if (e.type == 'drop') {
                this.files = e.dataTransfer.files;
                this.isDragIn = false;
            } else {
                this.files = e.target.files;
            }
            // 初始化
            this.startTime = new Date().getTime();
            this.logs = {};
            this.step = 0;
            if (this.files.length > 1) {
                this.showMultipleFileWarning = true;
            } else {
                this.showMultipleFileWarning = false;
            }
            // 所有文件并行处理
            for (var i = 0; i < this.files.length; i++) {
                handleFile(this.files[i]);
            }
        },
        log(key, value, variant) {
            // 用文件名做 key 在一些情况下会出 bug，比如要绑定到 v-b-toggle，空格和小数点这些都会影响正常绑定。
            // 所以先进行一次转码，页面上使用时再用 decodeURIComponent() 解码。
            key = encodeURIComponent(key);
            console.log(value);
            if (!this.logs[key]) {
                this.logs[key] = []
            }
            var timePast = '耗时 ' + (((new Date).getTime() - this.startTime) / 1000).toFixed(1) + 's';
            this.logs[key].push({
                timePast: timePast,
                value: value,
                variant: variant
            })
            // 由于vue限制，对于 logs 深层数值的多次更新，视图只显示初始化之后的第一次更新，
            // 因此通过更新位于数据顶层的进度条，即 data 里的 step，使深层日志的每次更新也都能被立即显示出来。
            this.step += 1
        }

    },
    computed: {
        selectAtLeastOne() {
            return this.selected.length > 0
        },
        selectedOptions() {
            // 数组转化为对象
            var selected = {}
            for (var i = 0; i < this.selected.length; i++) {
                selected[this.selected[i]] = true
            }
            return selected
        },
        selectedImageOptions() {
            // 数组转化为对象
            var selected = {}
            for (var i = 0; i < this.imagesOptionsSelected.length; i++) {
                selected[this.imagesOptionsSelected[i]] = true
            }
            return selected
        },
        computedSetting() {
            return {
                qualityTxt: Number(this.setting.quality).toFixed(2),
                sizeTxt: Number(this.setting.size).toFixed(2)
            }
        },
        totalStep() {
            return (this.files.length) * 8;
        }
    }

})

// 处理文件
async function handleFile(f) {
    app.log(f.name, '读取文件：' + f.name + ' 。');
    var zip = await JSZip.loadAsync(f).catch(function () {
        if (f.name.match(/\.xls$/)) {
            var msg = '解压错误，如果是旧版 xls 文件，请先在 Excel 中另存为 xlsx 格式。';
        } else {
            var msg = '解压错误，这可能不是一个 Excel 文件。';
        }
        app.log(f.name, msg, 'danger')
        app.step += 6;
    });

    if (!zip) { return }

    // 处理 styles.xml 文件
    if (app.selectedOptions.styles) {
        if (!(zip.file('xl/styles.xml'))) {
            if (app.selectedOptions.workbook || app.selectedOptions.png) {
                var msg = '提示：文档中没有 /xl/styles.xml，跳过。';
                app.log(f.name, msg, 'warning');
                app.step += 1;
            } else {
                var msg = '错误：文档中没有 /xl/styles.xml。如果是旧版 xls 文件，请先在 Excel 中另存为 xlsx 格式。';
                app.log(f.name, msg, 'danger');
                app.step += 6;
                return
            }
        } else {
            app.log(f.name, '读取内部 /xl/styles.xml 文件。');
            var stylesXML = await zip.file('xl/styles.xml').async('string');
            // 清除自定义样式
            app.log(f.name, '清除自定义样式。')
            // 不建议清除 numFmts 会造成一些日期、数字格式丢失。
            // stylesXML = stylesXML.replace(/<numFmts.*<\/numFmts>/, '');
            stylesXML = stylesXML.replace(/<cellStyleXfs.*<\/cellStyleXfs>/, '');
            stylesXML = stylesXML.replace(/<cellStyles.*<\/cellStyles>/, '');
            // 修改后的文件写入zip
            zip.file('xl/styles.xml', stylesXML);
        }
    } else {
        app.step += 2;
    }

    // 处理 workbook.xml 文件
    if (app.selectedOptions.workbook) {
        if (!(zip.file('xl/workbook.xml'))) {
            if (app.selectedOptions.png) {
                var msg = '提示：文档中没有 /xl/workbook.xml，跳过。';
                app.log(f.name, msg, 'warning');
                app.step += 1;
            } else {
                var msg = '错误：文档中没有 /xl/workbook.xml。如果是旧版 xls 文件，请先在 Excel 中另存为 xlsx 格式。';
                app.log(f.name, msg, 'danger');
                app.step += 4;
                return
            }
        } else {
            app.log(f.name, '读取内部 /xl/workbook.xml 文件。');
            var workbookXML = await zip.file('xl/workbook.xml').async('string');
            // 清除自定义样式
            app.log(f.name, '清除自定义名称。')
            workbookXML = workbookXML.replace(/<definedNames.*<\/definedNames>/, '');
            // 修改后的文件写入zip
            zip.file('xl/workbook.xml', workbookXML);
        }
    } else {
        app.step += 2;
    }

    // 处理 png 图片
    if (app.selectedOptions.png) {
        // 获取图片列表
        var imagesNameList = [];
        var imagesList = [];
        zip.forEach(function (relativePath) {
            if (relativePath.match(/\.(png|jpg|jpeg|webp)$/)) {
                imagesNameList.push(relativePath)
            }
        })
        if (!imagesNameList[0]) {
            if (app.selected.length < 2) {
                app.log(f.name, '错误：文档中没有图片，终止操作。如果是旧版 xls 文件，请先在 Excel 中另存为 xlsx 格式。', 'danger');
                app.step += 2;
                return
            } else {
                app.log(f.name, '提示：文档中没有图片，跳过。', 'warning');
            }
        } else {
            app.log(f.name, '正在处理图片，请耐心等候...');
            console.log('图片列表：', imagesNameList);
        }
        // 图片格式转换
        for (var i = 0; i < imagesNameList.length; i++) {
            imagesList[i] = await zip.file(imagesNameList[i]).async('blob');
            var base64 = await blob2Base64(imagesList[i]);
            var newImage = await base642Blob(base64, imagesNameList[i]);
            // 如果转换后比原文件还大，就不替换。
            if (newImage.size < imagesList[i].size) {
                app.log(f.name, '已替换 ' + imagesNameList[i] + '，' + (imagesList[i].size / 1024).toFixed(1) + 'KB → ' + (newImage.size / 1024).toFixed(1) + 'KB 。');
                imagesList[i] = newImage;
                app.step -= 1;
            } else {
                app.log(f.name, imagesNameList[i] + ' ' + (imagesList[i].size / 1024).toFixed(1) + 'KB → ' + (newImage.size / 1024).toFixed(1) + 'KB' + ' 转换后比原文件还大，不进行替换。', 'warning');
                app.step -= 1;
            }
            zip.file(imagesNameList[i], imagesList[i]);
        }

    } else {
        app.step += 2;
    }

    // 重新打包
    app.log(f.name, '正在重新压缩文件，请耐心等候...');
    zip.generateAsync({
        type: 'blob',
        compression: 'DEFLATE',
        compressionOptions: { level: 6 }
    }).then(function (blob) {
        // 输出到FileSaver.min.js
        saveAs(blob, f.name);
        var msg = (f.size / 1024).toFixed(1) + 'KB → ' + (blob.size / 1024).toFixed(1) + 'KB，压缩比例：' + (100 - blob.size / f.size * 100).toFixed(1) + '%。'
        app.log(f.name, '√ 完成！' + msg, 'success');
    });

}


// 以下函数复制自 https://github.com/renzhezhilu/webp2jpg-online 并做了一些修改
// 生成base64
function blob2Base64(blob) {
    return new Promise((resolve, reject) => {
        let reader = new FileReader()
        reader.readAsDataURL(blob)
        reader.onload = function () {
            resolve(this.result)
        }
    })
}
// base64还原成图片  type = 'jpeg/png/webp'  size 尺寸   quality 压缩质量
function base642Blob(base64, imageName) {
    return new Promise((resolve, reject) => {
        let imageType = 'image/jpeg'

        if (imageName.match(/\.webp$/)) {
            imageType = 'image/webp'
        }

        if (!app.selectedImageOptions.png2Jpeg && imageName.match(/\.png$/)) {
            imageType = 'image/png'
        }

        console.log(imageType)

        let size = Number(app.setting.size)
        let quality = Number(app.setting.quality)
        let img = new Image()
        img.src = base64
        img.onload = function () {
            // 只在内存中操作
            let _canvas = document.createElement('canvas')
            // 缩放
            let w = this.width * size
            let h = this.height * size
            _canvas.setAttribute("width", w)
            _canvas.setAttribute("height", h)
            _canvas.getContext("2d").drawImage(this, 0, 0, w, h)
            // 输出
            _canvas.toBlob(function (blob) {
                resolve(blob)
            }, imageType, quality)
        }
    })
}