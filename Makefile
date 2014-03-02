
.PHONY: docs test

all: test


test:
	prove --exec=lua --ext=lua -r test

test_all: test5.1 test5.2

test5.1:
	prove --exec=lua5.1 --ext=lua -r test

test5.2:
	prove --exec=lua5.2 --ext=lua -r test

test_travis:
	prove --exec="$LUA" --ext=lua -r test
