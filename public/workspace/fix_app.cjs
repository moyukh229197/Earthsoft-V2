const fs = require('fs');
const filePath = '/Users/moyukhroy/Desktop/My Software/Earthsoft/Earthsoft M 2/public/workspace/app.js';
let code = fs.readFileSync(filePath, 'utf8');

const target1 = `          if (i !== 0 || currentKm >= 0) {
            const canvas = document.createElement('canvas');
            canvas.width = 256;
            canvas.height = 128;
            const ctx = canvas.getContext('2d');
            ctx.fillStyle = '#facc15';
            ctx.fillRect(0, 0, 256, 128);
            ctx.fillStyle = '#0f172a';
            ctx.font = 'bold 54px sans-serif';
            ctx.textAlign = 'center';
            ctx.textBaseline = 'middle';
            ctx.fillText(\`KM \${currentKm}\`, 128, 64);
            ctx.lineWidth = 10;
            ctx.strokeStyle = '#0f172a';
            ctx.strokeRect(0, 0, 256, 128);
            const tex = new THREE.CanvasTexture(canvas);
            const board = new THREE.Mesh(new THREE.BoxGeometry(5, 3, 0.2), new THREE.MeshBasicMaterial({ map: tex }));
            board.position.copy(get3DViewerFramePoint(frame, centerOffset + (currentFormationW / 2) + 5, yPos + 5, 0));
            if (frame) board.rotation.y = frame.headingY;
            labelsGroup.add(board);
            const post = new THREE.Mesh(new THREE.BoxGeometry(0.4, 5, 0.4), new THREE.MeshBasicMaterial({ color: 0x334155 }));
            post.position.copy(get3DViewerFramePoint(frame, centerOffset + (currentFormationW / 2) + 5, yPos + 2.5, 0));
            if (frame) post.rotation.y = frame.headingY;
            labelsGroup.add(post);
            lastRenderedKm = currentKm;
          }`;

const replacement1 = `          if (i !== 0 || currentKm >= 0) {
            if (frame) {
              const kmStoneOffset = centerOffset - (currentFormationW / 2) - 3.5;
              const stone = createKMStone(frame, kmStoneOffset, yPos, \`KM \${currentKm}\`);
              labelsGroup.add(stone);
            }
            lastRenderedKm = currentKm;
          }`;

if(code.includes(target1)) {
    code = code.replace(target1, replacement1);
    console.log("Chunk 1 replaced successfully");
} else {
    console.log("Chunk 1 target not found");
}

fs.writeFileSync(filePath, code);
