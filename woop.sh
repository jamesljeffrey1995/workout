echo "Build number is $BUILD_NUMBER"

if [$BUILD_NUMBER % 2 == 0]
then
	python3 main.py
fi
