
.PHONY: docs test

all: test

test:
	@lua -v
	prove --exec=lua --ext=lua -r test

testall: test5.1 test5.2 testluajit

test5.1:
	@lua5.1 -v
	prove --exec=lua5.1 --ext=lua -r test

test5.2:
	@lua5.2 -v
	prove --exec=lua5.2 --ext=lua -r test

testluajit:
	@luajit -v
	prove --exec=luajit --ext=lua -r test
