# Master thesis analysis
This repository contains the R scripts used for the data analysis performed in my master thesis.

# Experiment
The scripts provided here were used for the analysis of the Ae200 titration experiment.
Other experiments were analyzed using the same script structure and workflow.

# Contents
The repository includes scripts for:
Data import and processing
Statistical analysis
Generation of plots

# Software
The analysis was performed using R Studio version 4.3.1

# Code origin
The scripts were originally created by a previous member of the lab, M.Sc. Inge Scharpf, with the assistance of OpenAI.

I adapted and modified these scripts for my specific analysis needs. For example, I recoded the script to generate spiderweb plots. The original version compared the mock control with a single treatment gourp. My version modified version compares all treatment groups with each other using different symbols and line types to distinguish the groups.

these modifications were implemented by me with the assistance of OpenAI.

# Usage of Codes

The scripts were used in the chronological oder S1-S5.
First each individual experiment was processed using S1, then all experimental replicates were processed using S2. Afterwards the statistical analysis was performed with the S3 script, followed by S4 and S5 to generate scatterplots and spiderweb plots.
