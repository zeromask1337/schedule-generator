SHAGGY_PATH=/Users/flippy/GolandProjects/awesomeProject

GOOS=linux GOARCH=amd64 go build -o $SHAGGY_PATH/bin/shaggy-amd64-linux $SHAGGY_PATH/src/*.go && GOOS=darwin GOARCH=amd64 go build -o $SHAGGY_PATH/bin/shaggy-amd64-darwin $SHAGGY_PATH/src/*.go