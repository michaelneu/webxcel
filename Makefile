xlsm: build
	cd src/xlsm && zip ../../build/webxcel.xlsm * -r

build:
	@mkdir build

clean:
	@rm -rf build