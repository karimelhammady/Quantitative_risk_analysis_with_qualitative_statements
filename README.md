# Quantitative_risk_analysis_with_qualitative_statements

## Description
 A spreadsheet-based tool that uses linguistic terms to estimate the probability of risk occurrence to help experts apply quantitative estimation easily by using common language as input, thus eliminating the need to assign precise probabilities. I have been working diligently on this project for the past year. 

 ## Requiremnts 

 1. [XLRISK](https://github.com/pyscripter/XLRisk) excel add-in 
 2. Latest excel version installed
 3. Enable macros on excel

 ## Usage 

### Steps 

1. Add tables in the config.xlsm file that maps the linguistic words to probability ranges. A sample of the tables is included in the config.xlsm file.
2. Adjust the parameters in the config.xlsm file before running the model. 
3. Open the main.xlsm file and add the risks along with the likelihood and loss for each risk. 
4. After defining the risks, press the run button to execute the model. Output folder will be created in the same running directory along with a Results.xlsm file. 


**Disclaimer:** This tool is currently under active development, and it may have bugs or incomplete features. We are working diligently to enhance its functionality and stability. Please use it with caution and report any issues you encounter.
