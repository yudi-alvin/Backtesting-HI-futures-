Please compile HI 1, HI 2, HI 3, HI 4 into one folder and name it 'HI'. 
All the indicators codes will read in the prices in the daily csv files and use the resampling method to compile the second (or micro second) ticks data into q-time data with OHLC (open, high, low, close).
All the analysis is done for intra-day, meaning I will close out the position by the end of each day and start a new position the next day (no overnight trading).<br/>
I had provided the following technical indicator code:<br/>
-Orderflow (OF1.py)<br/>
-Stochastic oscillator (Stochastic.py)<br/>
-Moving average (movingAverage.py)<br/>
you can use this code as a starting point and subsequently insert your own technical indicators for back-testing
