#copy the commands to Windows PowerShell and execute
SET GOOS=windows
SET GOARCH=386
go build -ldflags "-s -w" -buildmode=c-shared -o ../Release/excel-win-32.dll

SET GOOS=linux
SET GOARCH=386
go build -ldflags "-s -w" -buildmode=c-shared -o ../Release/excel-linux-32.so

SET GOOS=darwin
SET GOARCH=386
go build -ldflags "-s -w" -buildmode=c-shared -o ../Release/excel-darwin-32.so