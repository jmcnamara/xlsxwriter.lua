
.PHONY: docs test

all: test


test:
	prove --exec=lua --ext=lua -r test
