echo "Starting Installation"
sudo cp Webdrivers/chromedriver /usr/bin/
echo "Webdrivers copied"
sudo apt-get install xclip
echo 'Installed xclip for pyperclip module in python'
sudo apt-get install python3-venv
echo "VENV INSTALLED"
python3 -m venv venv
echo "VENV Created"
. venv/bin/activate
echo 'venv activated'
pip install -r req.txt
echo 'Requirements installed'
crontab -l > tmpfile
dir=`pwd`
echo "* * * * * ${dir}/run.sh">>tmpfile
crontab tmpfile
rm tmpfile
echo 'Cronjob was created successfully'
touch runLog.txt
touch debug.txt
sudo chmod +x run.sh
echo 'run.sh given executable permissions'
echo 'Initial Installation completed'
echo 'Ready to use =3'
