if (($BUILD_NUMBER%2==0));
then
    export $EMAIL_PASS
    python3 main.py
fi
