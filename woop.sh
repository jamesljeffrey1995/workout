echo "Build number is $BUILD_NUMBER"

if [$BUILD_NUMBER % 2 == 0]
	python3 main.py
fi
