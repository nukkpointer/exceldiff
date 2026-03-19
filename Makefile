build:
	go build -o exceldiff .

test:
	go test ./...

clean:
	rm -f exceldiff

.PHONY: build test clean
