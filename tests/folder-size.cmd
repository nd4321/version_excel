:: Attribution to: https://gist.github.com/sharunkumar/8caf903c146cbcc4c341571a6c6562df
@echo off
setlocal ENABLEDELAYEDEXPANSION
set firstline=.
set secondline=.
@FOR /F "tokens=1 delims=^^" %%G IN ('dir /s /a "%~1"') DO (
	set firstline=!secondline!
	set secondline=%%G
)
echo !firstline!