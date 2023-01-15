#!/bin/bash
clear

#sudo apt install sysbench

for f in {1..9}
do
	vcgencmd measure_temp
	sysbench cpu --cpu-max-prime=25000 --num-threads=4 run >/dev/null
done

vcgencmd measure_temp
