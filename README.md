<div id="top"></div>

## Table of Contents
1. [Background](#background)
2. [Implementation Details](#implementation-details)
3. [Usage](#usage)

## Background
This script determines the number of net new assets that Lower Mainland Biomedical Engineering (LMBME) receives and supports over 
a specified time period. It outputs a .xlsx file containing tabular data that can be used as input into the [LMBME service 
delivery cost model](https://github.com/jonathanjqchen/biomed-service-delivery-cost-model).

<p align="right"><a href="#top">[Back to top]</a></p>

## Implementation Details
The script determines net new assets by taking the difference in quantity between new assets and retired assets for the same time 
period. It also takes into consideration the following:

- Groups assets by model_num and site_code, which allows for determination of net new assets for each site/cost centre
- If an asset was retired but not replaced with a new asset of the same model at the same site, then it will show up in the final 
output with a negative quantity

Below is a control flow diagram outlining the implementation logic:

![net_new_script_control_flow](https://github.com/jonathanjqchen/biomed-net-new-assets-script/blob/main/images/control-flow-diagram.jpg)

<p align="right"><a href="#top">[Back to top]</a></p>

## Usage
To set up and use this project locally...

1. Install pandas and XlsxWriter
```
$ pip install pandas
$ pip install XlsxWriter
```

2. Clone the repo
```
$ git clone https://github.com/jonathanjqchen/biomed-net-new-assets-script.git
```

3. Follow the steps in ["README.pdf"](https://github.com/jonathanjqchen/biomed-net-new-assets-script/releases/tag/v1.0.0) on the releases page to export data on retired assets and new assets from TMS.

4. Run main.py 
```
$ cd biomed-net-new-assets-script
$ python main.py
```

5. Output will be in `./output/net_new_assets.xlsx`. There will be one worksheet in the file that looks like this:

![net_new_script_output](https://user-images.githubusercontent.com/54252001/147728265-226a23c5-93d4-483a-b665-a95e973ebadd.png)

<p align="right"><a href="#top">[Back to top]</a></p>
