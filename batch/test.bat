@echo off
setlocal EnableDelayedExpansion

set fruit=apple
echo The value of fruit is %fruit%

if "%fruit%"=="apple" (
  set fruit=orange
  echo The value of fruit within the block is !fruit!
  echo The value of fruit within the block is %fruit%
)

echo The value of fruit after the block is %fruit%