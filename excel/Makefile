BIN:=tempp

all: test build

build:
	go build -o ${BIN} *.go

fmt:
	go fmt

test:
	go test ./...

clean:
	rm -rf ${BIN}

