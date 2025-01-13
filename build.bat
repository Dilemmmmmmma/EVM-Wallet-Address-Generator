@echo off
echo 正在打包程序...
pyinstaller --clean wallet_generator.spec
echo 打包完成！
pause 