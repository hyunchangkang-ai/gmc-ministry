#!/bin/bash
cd /Users/hyunchangkang/Antigravity
echo "=== 1. 서버에서 최신 코드 가져오는 중 (Pull) ==="
git pull origin main
echo "=== 2. 변경된 내용 추가 중 ==="
git add .
echo "=== 3. 변경 사항 기록 중 ==="
git commit -m "Auto-sync from Mac: $(date +'%Y-%m-%d %H:%M:%S')"
echo "=== 4. 서버에 업로드 중 (Push) ==="
git push origin main
echo "=== 동기화가 완료되었습니다! ==="
