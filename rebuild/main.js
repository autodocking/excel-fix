
var app = new Vue({
    el: '#app',
    data: {
        alert: '读取错误',
        names: [],
        isDragIn: false,
    },
    mounted() {
    },
    methods: {
        onDrop(e) {
            console.log(e);
            this.isDragIn = false;
        },
        onDragenter(e) {
            this.isDragIn = true;
        },
        onDragleave(e) {
            this.isDragIn = false;
        },
    }

})

