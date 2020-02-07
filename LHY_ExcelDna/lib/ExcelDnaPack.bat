set packpath=%~pd0
set packname=%1
cd /d %packpath%
ExcelDnaPack.exe "%packname%.dna" /Y /O "%packname%Pack.xll"
ExcelDnaPack.exe "%packname%64.dna" /Y /O "%packname%Pack64.xll"
