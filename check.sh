while true
do
rm info.txt
rm wget-log
wget --no-check-certificate https://guyutongxue.github.io/2018_new_year_party/next/info.txt
cat info.txt
read
done