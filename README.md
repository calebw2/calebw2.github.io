# Random Forest Regression for the Manufacturing of Industrial Lubricating Grease


For the manufacturing of lubricating grease, one of the most critical specifications is the consistency of the grease. Often for the manufacturer, this is a difficult task, as the outcome does not always follow a logical progression, and often much time is wasted on incremental adjustments to achieve the desired consistency. The goal of this model is to assist the manufacturer in efficiency by predicting the approximate amount of base oil needed to meet specifications.

## Overview of Manufacturing Grease
Grease is defined as a lubricant suspended in some thickener. This thickener is most often a soap, the metallic salt of a fatty acid. There are several types of thickeners and complexing agents used for this. The general method for manufacture is as follows: 

1.	A “base grease” is cooked in a reactor by heating base oil, dissolving a fatty acid, adding a metallic base (i.e. lithium hydroxide), and heating until the reaction is complete
2.	The base grease is transferred to a finishing kettle where it receives its additives. The additive package is different for every grease according to its needs
3.	The consistency is tested, and the grease is then “oiled back” by adding more base oil, causing the grease to thin out. This step is done incrementally until the desired consistency is reached.
4.	The grease is filled into its packaging and shipped.



## The Problem
Through this relatively simple process, one usually reaches the conclusion that more oil yields a softer grease. While this is generally true, it isn't easy to reliably estimate how much oil will end up being needed to achieve the desired consistency. Not every grease type reacts the same, and sometimes even the same grease will respond differently. With much hands-on experience, the technician can sort of “get a feel” for it, but this requires a lot of commitment that is pretty unreasonable for most people in this position to put in given their compensation. As the firm grows and more people are involved in the process with less depth this becomes even more true.

There have been several attempts between myself and my coworkers to develop linear relationships, but all of these attempts have fallen short of being more accurate than intuition, let alone able to handle more than the specific product it was made for. It is clear that there are variables at play that have a non-negligible effect that would be a waste of time to try to account for mentally for each specific product, kettle, soap, etc. We have long suspected that machine learning would be critical in solving this issue and increasing efficiency; however, the resources and knowledge to implement this have not yet been in our grasp. I enrolled in this course for that very reason. 

## The Data

### Acquiring the Data

By far the most difficult part of this project has been sourcing and cleaning the data required to train a model. The data we have is fairly low quality given that only in recent years have we started rigorously collecting data in preparation for this type of project. Because of this, there have been many changes in formatting on the Excel sheets where we record. Also, all the data being stored in excel adds another layer in complexity, especially since I have little experience with Python. However I was eventually able to write a fairly rudimentary function that will scrape as much of our QC Analysis data as possible using openpyxl and collect it in another excel worksheet. The code for this can be found with this link. 

After this long process with many iterations, I was able to collect a few thousand batches. Some of the cleaning and filtering was beyond my knowledge of my Python capabilities, so the rest was sorted semi-manually using excel. After this even longer process I was able to collect complete data for 862 batches. Some values needed to be scaled and adjusted to be useful, and this was done mostly using excel. The final data set that was used to train and test the model can be found [here](/assets/final_matrix.csv).

### Overview and Analysis of the Data

In the final dataset, there are 22 columns of data, however not all are needed for the model. The product name, customer name, well load, and lot number are irrelevant factors so they were dropped. Furthermore, some values needed to be adjusted since the target variable is known in the training data but will not be known during practical use, so this must be accounted for. Several other categories were dropped because they were only there as indicators for me.

The date was changed to represent the integer number of day of the year using excel in order to be usable for the model. Including the date accounts for ambient temperature changes throughout the year, which seasonally is significant. The string version of the date was dropped.

Furthermore, the kettle that the batch is in as well as the thickener type used were given integer values as well, and the string versions of these data were also dropped; however, this was done in code using pandas after the final table was read, and the values simply replaced in each respective category.

The final dataset fed into the model consisted of 862 samples with 11 columns. The target variable of total base oil percentage was separated. This left the data with ten parameters to be considered:

•	date: in integer form to account for seasonal ambient temperature conditions

•	kettle: which of the nine kettles was used to process the batch

•	thickener type: in integer form, accounts for how different soaps respond to oilback

•	elco complex(?):  A boolean indicating whether a certain complexing agent was used that, in my experience, makes a significant difference in oil back response

•	base grease percentage (pre oil)

•	liquid additives percentage (pre-oil)

•	solid additives percentage (pre-oil)

•	silica percentage (pre-oil) – silica has a much more pronounced thickening effect per percentage point than any other additive. 

•	Rework percentage- accounts for “rework”, which refers to adding finished product that was not shipped for any number of reasons. This was at first omitted but added significant enough precision to warrant inclusion. 

## The Model

My first intuition was to use a ridge regression algorithm to find some linear relationship
