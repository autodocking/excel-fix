var $result = $('#result');
var dateBefore = new Date();
var timecounts = {}

$('#file').on('change', sizingXlsx);
$('body').on('dragover', function (e) { e.preventDefault(); });
$('body').on('drop', function (e) { e.preventDefault(); });

$('#dropZone').on('drop', sizingXlsx);
$('#dropZone').on('dragenter', function () { $('#dropZone').addClass('drag-over') });
$('#dropZone').on('dragleave', function () { $('#dropZone').removeClass('drag-over') });
$('#dropZone').on('drop', function () { $('#dropZone').removeClass('drag-over') });

function sizingXlsx(evt) {
    // 清空处理结果区域
    $result.html('');
    $('#result_block').removeClass('hidden').addClass('show');

    if (evt.type == 'drop') {
        var files = evt.originalEvent.dataTransfer.files;
    } else {
        var files = evt.target.files;
    }

    for (var i = 0; i < files.length; i++) {
        handleFile(files[i]);
    }

    function handleFile(f) {
        // 为每一个文件独立计时
        var pid = Math.random();
        timecounts[pid] = {}
        timecounts[pid].dateBefore = new Date();

        var $title = $('<h4>', {
            text: f.name
        });
        var $fileContent = $('<ul>');
        $result.append($title);
        $result.append($fileContent);

        JSZip.loadAsync(f)
            .then(function (zip) {
                log('已完成解压，请耐心等候...');
                if (!(zip.file('xl/styles.xml'))) {
                    log('找不到【 ./xl/styles.xml 】文件。', 1)
                    $result.append($('<div>', {
                        'class': 'alert alert-danger',
                        text: '读取错误：请确认文件【 ' + f.name + ' 】是否为xlsx格式，如果是旧版本的xls文件，请先在Excel中另存为xlsx格式。'
                    }));
                    return;
                }
                zip.file('xl/styles.xml').async('string').then(function (temp) {
                    // 清除自定义样式
                    temp = temp.replace(/<cellStyles.*<\/cellStyles>/, '')
                    log('已清除自定义样式。');
                    // 写入zip
                    zip.file('xl/styles.xml', temp);
                    // 输出zip到FileSaver.min.js
                    log('正在合成文件...', 1);
                    zip.generateAsync({
                        type: 'blob',
                        compression: 'DEFLATE',
                        compressionOptions: { level: 6 }
                    })
                        .then(function (blob) {
                            log('已完成！请留意下载内容。');
                            // 开始下载
                            saveAs(blob, f.name);
                            log('请在Excel中打开并重新保存一次该文件，将进一步缩小体积。', 1);
                        });
                });
            }, function (e) {
                $result.append($('<div>', {
                    'class': 'alert alert-danger',
                    text: '读取错误：请确认文件【 ' + f.name + ' 】是否为xlsx格式，如果是旧版本的xls文件，请先在Excel中另存为xlsx格式。'
                }));
                console.log(f.name + ': ' + e.message);
            });

        function log(text, a) {
            if (a == undefined) {
                $fileContent.append($("<li>", {
                    text: text + '(' + timePast() + 'ms)'
                }));
            } else {
                $fileContent.append($("<li>", {
                    text: text
                }));
            }
        }

        function timePast() {
            timecounts[pid].dateAfter = new Date();
            var timePast = timecounts[pid].dateAfter - timecounts[pid].dateBefore;
            timecounts[pid].dateBefore = timecounts[pid].dateAfter;
            return timePast;
        }
    }
}