@echo off
for /f "tokens=*" %%A in (commands.txt) do (
    echo Executing: %%A
    %%A
)
