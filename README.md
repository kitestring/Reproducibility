# Reproducibility_08-15-2011

Date Written: 08/15/2011

Industry: Time of Flight Mass Spectrometer Developer & Manufacturer

Department: Hardware & Software Customer Support

GUI: “GUI_AnalytesFields_Tab.png” & “GUI_Instrumentation_Tab.png”

Sample Raw Data:

“GROB Z-Test EG1@1600_60.csv” This *.csv file contains the processed data exported from the software interfaced to our Time of Flight Mass Spectrometer.  Within this file is every chemical the instrument found within a single sample along with each user defined metric.  Every sample measurement would generate a single *.csv file.  For a single experiment it was not uncommon to generate hundreds of *.csv files.

Sample Output:

Only a tiny fraction of the data can reasonably be shown here due to the sheer volume that is generated.  To summarize… 
For each chemical included in the analysis an excel worksheet is generated.  It contains all of the data from each field (metric) in tabular form as well as shown graphically (“SampleOutput_IndividualAnalyteTable.png” & “SampleOutput_IndividualAnalytePlot.png”).  Additionally, there is also a summary worksheet generated that calculates an average, standard deviation, and relative standard deviation for each metric (“SampleOutput_SummaryTable.png “).  The summary worksheet also plots each metric, with all chemicals on a single plot (“SampleOutput_SummaryPlot.png”).
