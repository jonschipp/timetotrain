cleanup:
	$(info [$(YELLOW)*$(NORMAL)] Removing spreadsheets)
	rm -f *.xlsx

install:
	$(info [$(YELLOW)*$(NORMAL)] Installing dependencies)
	python3 setup.py build
	python3 setup.py install
