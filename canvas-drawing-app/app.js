const canvas = document.getElementById('drawingCanvas');
const ctx = canvas.getContext('2d');
const colorPicker = document.getElementById('colorPicker');
const sizePicker = document.getElementById('sizePicker');
const clearBtn = document.getElementById('clearBtn');
const saveBtn = document.getElementById('saveBtn');

let isDrawing = false;
let lastX = 0;
let lastY = 0;

// 初始化畫布背景為白色
ctx.fillStyle = 'white';
ctx.fillRect(0, 0, canvas.width, canvas.height);

// 繪圖設定
ctx.lineJoin = 'round';
ctx.lineCap = 'round';

function draw(e) {
    if (!isDrawing) return;

    // 取得當前滑鼠位置
    const x = e.offsetX;
    const y = e.offsetY;

    // 設定畫筆樣式
    ctx.strokeStyle = colorPicker.value;
    ctx.lineWidth = sizePicker.value;

    // 繪製線條
    ctx.beginPath();
    ctx.moveTo(lastX, lastY);
    ctx.lineTo(x, y);
    ctx.stroke();

    // 更新最後位置
    lastX = x;
    lastY = y;
}

// 綁定滑鼠事件
canvas.addEventListener('mousedown', (e) => {
    isDrawing = true;
    lastX = e.offsetX;
    lastY = e.offsetY;
});

canvas.addEventListener('mousemove', draw);
canvas.addEventListener('mouseup', () => isDrawing = false);
canvas.addEventListener('mouseout', () => isDrawing = false);

// 清除畫布
clearBtn.addEventListener('click', () => {
    ctx.fillStyle = 'white';
    ctx.fillRect(0, 0, canvas.width, canvas.height);
});

// 儲存圖片
saveBtn.addEventListener('click', () => {
    const link = document.createElement('a');
    link.download = 'my-drawing.png';
    link.href = canvas.toDataURL('image/png');
    link.click();
});
