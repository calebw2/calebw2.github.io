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

There have been several attempts between myself and my coworkers to develop linear relationships, but all of these attempts have fallen short of being more accurate than intuition, let alone able to handle more than the specific product it was made for. We have long suspected that machine learning would be critical in solving this issue and increasing efficiency, however, the resources and knowledge to implement this have not yet been in our grasp. I enrolled in this course for that very reason. 


