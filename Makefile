xlsm: build
	cd src && zip ../build/webxcel.xlsm * -r

build:
	@mkdir build

clean:
	@rm -rf build