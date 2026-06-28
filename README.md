# pdftoexcel
bash install.sh
LOAD DATA LOCAL INFILE '/data/data/com.termux/files/home/pdftoexel/produk.csv'
INTO TABLE produk
FIELDS TERMINATED BY ','
OPTIONALLY ENCLOSED BY '"'
LINES TERMINATED BY '\n'
IGNORE 1 LINES;
