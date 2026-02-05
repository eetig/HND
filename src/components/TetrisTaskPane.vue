<template>
  <div class="tetris-container">
    <div class="game-header">
      <div class="game-controls">
        <button @click="startGame" :disabled="isPlaying" class="btn-start">开始</button>
        <button @click="pauseGame" :disabled="!isPlaying" class="btn-pause">暂停</button>
        <button @click="resetGame" class="btn-reset">重置</button>
      </div>
    </div>
    <div class="game-content">
      <div class="game-board">
        <canvas ref="gameCanvas" width="240" height="360"></canvas>
      </div>
      <div class="game-info">
        <div class="info-item">
          <h3>得分</h3>
          <div class="score">{{ score }}</div>
        </div>
        <div class="info-item">
          <h3>等级</h3>
          <div class="level">{{ level }}</div>
        </div>
        <div class="info-item">
          <h3>行数</h3>
          <div class="lines">{{ linesCleared }}</div>
        </div>
        <div class="info-item">
          <h3>下一个</h3>
          <canvas ref="nextCanvas" width="80" height="80"></canvas>
        </div>
        <div class="info-item">
          <h3>最高分</h3>
          <div class="high-score">{{ highScore }}</div>
        </div>
      </div>
    </div>
    <div class="game-footer">
      <div class="controls-info">
        <p>A ← 左移</p>
        <p>D → 右移</p>
        <p>W 旋转</p>
        <p>S 加速</p>
        <p>空格 暂停/继续</p>
      </div>
    </div>
    <div v-if="isGameOver" class="game-over">
      <h2>游戏结束</h2>
      <p>最终得分: {{ score }}</p>
      <button @click="resetGame" class="btn-restart">重新开始</button>
    </div>
    <div v-if="isPaused && isPlaying" class="game-paused">
      <h2>游戏暂停</h2>
      <p>按空格键继续游戏</p>
    </div>
  </div>
</template>

<script>
export default {
  name: 'TetrisTaskPane',
  data() {
    return {
      // 游戏状态
      isPlaying: false,
      isPaused: false,
      isGameOver: false,
      score: 0,
      level: 1,
      linesCleared: 0,
      highScore: 0,
      
      // 游戏配置
      boardWidth: 20,
      boardHeight: 30,
      cellSize: 12,
      
      // 游戏元素
      board: [],
      currentPiece: null,
      nextPiece: null,
      dropInterval: 1000,
      lastDropTime: 0,
      
      // Canvas上下文
      gameCtx: null,
      nextCtx: null,
      
      // 动画帧ID
      animationId: null
    }
  },
  mounted() {
    this.initCanvas();
    this.initBoard();
    this.loadHighScore();
    this.generateNextPiece();
    this.spawnPiece();
    this.render();
    this.setupKeyboardControls();
  },
  beforeUnmount() {
    this.stopGame();
    window.removeEventListener('keydown', this.handleKeyDown);
  },
  methods: {
    // 初始化Canvas
    initCanvas() {
      const gameCanvas = this.$refs.gameCanvas;
      const nextCanvas = this.$refs.nextCanvas;
      
      this.gameCtx = gameCanvas.getContext('2d');
      this.nextCtx = nextCanvas.getContext('2d');
    },
    
    // 初始化游戏板
    initBoard() {
      this.board = Array(this.boardHeight).fill().map(() => Array(this.boardWidth).fill(0));
    },
    
    // 加载最高分
    loadHighScore() {
      const savedScore = localStorage.getItem('tetris_high_score');
      this.highScore = savedScore ? parseInt(savedScore) : 0;
    },
    
    // 保存最高分
    saveHighScore() {
      if (this.score > this.highScore) {
        this.highScore = this.score;
        localStorage.setItem('tetris_high_score', this.highScore.toString());
      }
    },
    
    // 方块定义
    getTetrominos() {
      return [
        // I形
        {
          shape: [[0, 0, 0, 0], [1, 1, 1, 1], [0, 0, 0, 0], [0, 0, 0, 0]],
          color: '#888888'
        },
        // J形
        {
          shape: [[1, 0, 0], [1, 1, 1], [0, 0, 0]],
          color: '#666666'
        },
        // L形
        {
          shape: [[0, 0, 1], [1, 1, 1], [0, 0, 0]],
          color: '#999999'
        },
        // O形
        {
          shape: [[1, 1], [1, 1]],
          color: '#777777'
        },
        // S形
        {
          shape: [[0, 1, 1], [1, 1, 0], [0, 0, 0]],
          color: '#555555'
        },
        // T形
        {
          shape: [[0, 1, 0], [1, 1, 1], [0, 0, 0]],
          color: '#aaaaaa'
        },
        // Z形
        {
          shape: [[1, 1, 0], [0, 1, 1], [0, 0, 0]],
          color: '#444444'
        }
      ];
    },
    
    // 随机生成方块
    generatePiece() {
      const tetrominos = this.getTetrominos();
      const randomIndex = Math.floor(Math.random() * tetrominos.length);
      return tetrominos[randomIndex];
    },
    
    // 生成下一个方块
    generateNextPiece() {
      this.nextPiece = this.generatePiece();
      this.renderNextPiece();
    },
    
    // 生成新方块
    spawnPiece() {
      this.currentPiece = {
        tetromino: this.nextPiece,
        position: { x: Math.floor(this.boardWidth / 2) - 1, y: 0 }
      };
      this.generateNextPiece();
      
      // 检查游戏是否结束
      if (this.checkCollision(this.currentPiece.tetromino.shape, this.currentPiece.position)) {
        this.gameOver();
      }
    },
    
    // 检查碰撞
    checkCollision(shape, position) {
      for (let y = 0; y < shape.length; y++) {
        for (let x = 0; x < shape[y].length; x++) {
          if (shape[y][x]) {
            const newX = position.x + x;
            const newY = position.y + y;
            
            if (
              newX < 0 || 
              newX >= this.boardWidth || 
              newY >= this.boardHeight ||
              (newY >= 0 && this.board[newY][newX])
            ) {
              return true;
            }
          }
        }
      }
      return false;
    },
    
    // 旋转方块
    rotatePiece() {
      const rotatedShape = this.rotateMatrix(this.currentPiece.tetromino.shape);
      if (!this.checkCollision(rotatedShape, this.currentPiece.position)) {
        this.currentPiece.tetromino.shape = rotatedShape;
      }
    },
    
    // 旋转矩阵
    rotateMatrix(matrix) {
      const N = matrix.length;
      const result = Array(N).fill().map(() => Array(N).fill(0));
      
      for (let y = 0; y < N; y++) {
        for (let x = 0; x < N; x++) {
          result[x][N - 1 - y] = matrix[y][x];
        }
      }
      
      return result;
    },
    
    // 移动方块
    movePiece(direction) {
      const newPosition = { ...this.currentPiece.position };
      
      if (direction === 'left') {
        newPosition.x -= 1;
      } else if (direction === 'right') {
        newPosition.x += 1;
      } else if (direction === 'down') {
        newPosition.y += 1;
      }
      
      if (!this.checkCollision(this.currentPiece.tetromino.shape, newPosition)) {
        this.currentPiece.position = newPosition;
        return true;
      }
      
      return false;
    },
    
    // 硬降
    hardDrop() {
      while (this.movePiece('down')) {}
      this.lockPiece();
    },
    
    // 锁定方块
    lockPiece() {
      const { shape } = this.currentPiece.tetromino;
      const { x, y } = this.currentPiece.position;
      
      for (let i = 0; i < shape.length; i++) {
        for (let j = 0; j < shape[i].length; j++) {
          if (shape[i][j]) {
            const boardY = y + i;
            const boardX = x + j;
            
            if (boardY >= 0) {
              this.board[boardY][boardX] = this.currentPiece.tetromino.color;
            }
          }
        }
      }
      
      this.clearLines();
      this.spawnPiece();
    },
    
    // 清除行
    clearLines() {
      let linesCleared = 0;
      
      for (let y = this.boardHeight - 1; y >= 0; y--) {
        if (this.board[y].every(cell => cell)) {
          this.board.splice(y, 1);
          this.board.unshift(Array(this.boardWidth).fill(0));
          linesCleared++;
          y++; // 重新检查当前行
        }
      }
      
      if (linesCleared > 0) {
        this.updateScore(linesCleared);
        this.linesCleared += linesCleared;
        this.updateLevel();
      }
    },
    
    // 更新得分
    updateScore(linesCleared) {
      const points = [0, 100, 300, 500, 800];
      this.score += points[linesCleared] * this.level;
      this.saveHighScore();
    },
    
    // 更新等级
    updateLevel() {
      const newLevel = Math.floor(this.linesCleared / 10) + 1;
      if (newLevel > this.level) {
        this.level = newLevel;
        this.dropInterval = 1000 - (this.level - 1) * 100;
        if (this.dropInterval < 100) {
          this.dropInterval = 100;
        }
      }
    },
    
    // 游戏结束
    gameOver() {
      this.isPlaying = false;
      this.isGameOver = true;
      this.stopGameLoop();
    },
    
    // 开始游戏
    startGame() {
      if (!this.isPlaying && !this.isGameOver) {
        this.isPlaying = true;
        this.isPaused = false;
        this.lastDropTime = Date.now();
        this.gameLoop();
      }
    },
    
    // 暂停游戏
    pauseGame() {
      this.isPaused = !this.isPaused;
      if (!this.isPaused) {
        this.lastDropTime = Date.now();
        this.gameLoop();
      } else {
        this.stopGameLoop();
      }
    },
    
    // 重置游戏
    resetGame() {
      this.stopGameLoop();
      this.initBoard();
      this.score = 0;
      this.level = 1;
      this.linesCleared = 0;
      this.isPlaying = false;
      this.isPaused = false;
      this.isGameOver = false;
      this.dropInterval = 1000;
      this.spawnPiece();
      this.render();
    },
    
    // 停止游戏
    stopGame() {
      this.stopGameLoop();
      this.isPlaying = false;
    },
    
    // 游戏循环
    gameLoop() {
      if (!this.isPlaying || this.isPaused) return;
      
      const now = Date.now();
      if (now - this.lastDropTime > this.dropInterval) {
        if (!this.movePiece('down')) {
          this.lockPiece();
        }
        this.lastDropTime = now;
      }
      
      this.render();
      this.animationId = requestAnimationFrame(this.gameLoop);
    },
    
    // 停止游戏循环
    stopGameLoop() {
      if (this.animationId) {
        cancelAnimationFrame(this.animationId);
        this.animationId = null;
      }
    },
    
    // 渲染游戏
    render() {
      const ctx = this.gameCtx;
      if (!ctx) return;
      
      // 清空画布
      ctx.clearRect(0, 0, 240, 360);
      
      // 绘制网格
      ctx.strokeStyle = '#888';
      ctx.lineWidth = 1;
      
      for (let x = 0; x <= this.boardWidth; x++) {
        ctx.beginPath();
        ctx.moveTo(x * this.cellSize, 0);
        ctx.lineTo(x * this.cellSize, 360);
        ctx.stroke();
      }
      
      for (let y = 0; y <= this.boardHeight; y++) {
        ctx.beginPath();
        ctx.moveTo(0, y * this.cellSize);
        ctx.lineTo(240, y * this.cellSize);
        ctx.stroke();
      }
      
      // 绘制已锁定的方块
      for (let y = 0; y < this.boardHeight; y++) {
        for (let x = 0; x < this.boardWidth; x++) {
          if (this.board[y][x]) {
            this.drawCell(ctx, x, y, this.board[y][x]);
          }
        }
      }
      
      // 绘制当前方块
      if (this.currentPiece) {
        const { shape, color } = this.currentPiece.tetromino;
        const { x, y } = this.currentPiece.position;
        
        for (let pieceY = 0; pieceY < shape.length; pieceY++) {
          for (let pieceX = 0; pieceX < shape[pieceY].length; pieceX++) {
            if (shape[pieceY][pieceX]) {
              const boardX = x + pieceX;
              const boardY = y + pieceY;
              
              if (boardY >= 0) {
                this.drawCell(ctx, boardX, boardY, color);
              }
            }
          }
        }
      }
    },
    
    // 渲染下一个方块
    renderNextPiece() {
      const ctx = this.nextCanvas;
      if (!ctx || !this.nextPiece) return;
      
      // 清空画布
      ctx.clearRect(0, 0, 80, 80);
      
      // 绘制下一个方块
      const { shape, color } = this.nextPiece;
      const offsetX = (80 - shape.length * 20) / 2;
      const offsetY = (80 - shape.length * 20) / 2;
      
      for (let y = 0; y < shape.length; y++) {
        for (let x = 0; x < shape[y].length; x++) {
          if (shape[y][x]) {
            ctx.fillStyle = color;
            ctx.fillRect(offsetX + x * 20, offsetY + y * 20, 19, 19);
            ctx.strokeStyle = '#000';
            ctx.lineWidth = 1;
            ctx.strokeRect(offsetX + x * 20, offsetY + y * 20, 19, 19);
          }
        }
      }
    },
    
    // 绘制单元格
    drawCell(ctx, x, y, color) {
      ctx.fillStyle = color;
      ctx.fillRect(x * this.cellSize, y * this.cellSize, this.cellSize - 1, this.cellSize - 1);
      ctx.strokeStyle = '#000';
      ctx.lineWidth = 1;
      ctx.strokeRect(x * this.cellSize, y * this.cellSize, this.cellSize - 1, this.cellSize - 1);
    },
    
    // 键盘控制
    setupKeyboardControls() {
      window.addEventListener('keydown', this.handleKeyDown);
    },
    
    // 处理键盘事件
    handleKeyDown(event) {
      if (!this.isPlaying || this.isPaused) {
        if (event.code === 'Space') {
          this.pauseGame();
        }
        return;
      }
      
      switch (event.code) {
        case 'KeyA':
          this.movePiece('left');
          break;
        case 'KeyD':
          this.movePiece('right');
          break;
        case 'KeyS':
          this.movePiece('down');
          break;
        case 'KeyW':
          this.rotatePiece();
          break;
        case 'Space':
          this.pauseGame();
          break;
      }
    }
  }
}
</script>

<style scoped>
.tetris-container {
  width: 100%;
  height: 100%;
  padding: 15px;
  background-color: #f5f5f5;
  color: #333333;
  font-family: 'Arial', sans-serif;
}

.game-header {
  text-align: center;
  margin-bottom: 15px;
}

.game-header h1 {
  font-size: 20px;
  margin-bottom: 10px;
  color: #333333;
  font-weight: bold;
}

.game-controls {
  display: flex;
  justify-content: center;
  gap: 8px;
}

.game-controls button {
  padding: 6px 12px;
  border: 1px solid #999999;
  border-radius: 3px;
  cursor: pointer;
  font-weight: bold;
  background-color: #ffffff;
  color: #333333;
  transition: all 0.3s;
  box-shadow: 1px 1px 2px rgba(0, 0, 0, 0.1);
}

.game-controls button:hover {
  background-color: #e0e0e0;
  box-shadow: 1px 1px 3px rgba(0, 0, 0, 0.2);
}

.game-controls button:disabled {
  background-color: #f0f0f0;
  color: #999999;
  cursor: not-allowed;
  box-shadow: none;
}

.game-content {
  display: flex;
  justify-content: center;
  gap: 20px;
  margin-bottom: 15px;
}

.game-board {
  border: 2px solid #999999;
  border-radius: 3px;
  background-color: #ffffff;
  box-shadow: 2px 2px 4px rgba(0, 0, 0, 0.1);
}

canvas {
  display: block;
}

.game-info {
  display: flex;
  flex-direction: column;
  gap: 10px;
  min-width: 100px;
  max-width: 120px;
}

.info-item {
  background-color: #ffffff;
  padding: 8px;
  border-radius: 3px;
  text-align: center;
  border: 1px solid #dddddd;
  box-shadow: 1px 1px 2px rgba(0, 0, 0, 0.05);
}

.info-item h3 {
  font-size: 12px;
  margin: 0 0 3px 0;
  color: #666666;
  font-weight: bold;
}

.score, .level, .lines, .high-score {
  font-size: 16px;
  font-weight: bold;
  color: #333333;
}

#nextCanvas {
  margin: 0 auto;
  background-color: #ffffff;
  border: 1px solid #dddddd;
}

.game-footer {
  margin-top: 15px;
  text-align: center;
}

.controls-info {
  display: flex;
  justify-content: center;
  gap: 15px;
  flex-wrap: wrap;
}

.controls-info p {
  margin: 0;
  font-size: 11px;
  color: #666666;
}

.game-over {
  position: absolute;
  top: 50%;
  left: 50%;
  transform: translate(-50%, -50%);
  background-color: rgba(255, 255, 255, 0.95);
  padding: 20px;
  border-radius: 6px;
  text-align: center;
  z-index: 1000;
  border: 1px solid #dddddd;
  box-shadow: 2px 2px 10px rgba(0, 0, 0, 0.1);
}

.game-over h2 {
  margin-top: 0;
  color: #666666;
  font-size: 18px;
}

.btn-restart {
  margin-top: 15px;
  padding: 8px 16px;
  background-color: #ffffff;
  color: #333333;
  border: 1px solid #999999;
  border-radius: 3px;
  cursor: pointer;
  font-weight: bold;
  box-shadow: 1px 1px 2px rgba(0, 0, 0, 0.1);
}

.btn-restart:hover {
  background-color: #e0e0e0;
}

.game-paused {
  position: absolute;
  top: 50%;
  left: 50%;
  transform: translate(-50%, -50%);
  background-color: rgba(255, 255, 255, 0.95);
  padding: 20px;
  border-radius: 6px;
  text-align: center;
  z-index: 1000;
  border: 1px solid #dddddd;
  box-shadow: 2px 2px 10px rgba(0, 0, 0, 0.1);
}

.game-paused h2 {
  margin-top: 0;
  color: #666666;
  font-size: 18px;
}

/* 响应式设计 */
@media (max-width: 768px) {
  .game-content {
    flex-direction: row;
    align-items: flex-start;
  }
  
  .game-info {
    flex-direction: column;
    flex-wrap: nowrap;
    justify-content: flex-start;
  }
  
  .info-item {
    min-width: 90px;
  }
}
</style>