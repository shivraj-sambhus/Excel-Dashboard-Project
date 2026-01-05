# Excel-Dashboard-Project
<figure>
  <img width="1298" height="557" alt="image" src="https://github.com/user-attachments/assets/1de010dd-3700-4521-8f6a-b29734677027" />
  <figcaption>      Total coffee sales from 2019 to 2022</figcaption>
</figure>

<img width="1088" height="99" alt="image" src="https://github.com/user-attachments/assets/a35c727d-e45b-4445-83da-c07e4b516a4b" />


<figure>
  <img width="1298" height="557" alt="image" src="https://github.com/user-attachments/assets/f622fbc0-de6e-4243-a5da-e8bc0797b7fe" />
  <figcaption>    Total coffee sales of medium roast blends purchased by customers without loyalty cards from 2019 to 2022</figcaption>
</figure>

<img width="967" height="40" alt="image" src="https://github.com/user-attachments/assets/5adc264d-9d24-4eaf-bbd7-e298ec85b607" />

# Objective

In this project, I will create an interactive dashboard of coffee sales that can be filtered by year, quantity sold, roast type, and by members of the loyalty program. This allows me to analyze the customer sales dataset and understand the frequency of purchases by customer and country. Note that this project was completed in Google Sheets and you may encounter issues with compiling it if you use Excel directly.

<img width="967" height="40" alt="image" src="https://github.com/user-attachments/assets/c6ab204a-bb10-4245-ab20-1771c4897068" />



# Data

The provided Excel file includes three structured datasets: orders, customers, and products. The orders dataset includes details of individual coffee orders, including the customer ID, the order date, customer emails, and so on. It has 1001 rows and 13 columns.  The customers dataset includes details of each customer, such as their name, address, loyalty card information, and so on. It has 1001 rows and 8 columns. The products dataset includes details of the coffee sold, such as the name of the roast, the quantity produced, the selling price, and so on. It has 49 rows and 7 columns and contains coffee data from 2019 to 2022 in the U.S.A, U.K, and Ireland.

<img width="967" height="40" alt="image" src="https://github.com/user-attachments/assets/a4d86e7a-348a-4cbd-85c1-06918357a6fb" />


# Methods

I began my project with the orders worksheet. The worksheet originally had no entries in the Customer Name, Email, Country, Coffee Type, Roast Type, Size, Unit Price, and Sales columns. Therefore, I started by populating the Customer Name column with the XLOOKUP formula below. I used a similar formula to populate the Email and Country columns.

<img width="639" height="30" alt="image" src="https://github.com/user-attachments/assets/29637fec-b43d-4e0e-9c8f-68e4ec0b22cc" />



Rather than repeating the XLOOKUP formula several more times, I used the INDEX(MATCH(MATCH)) formula below to populate the Coffee Type, Roast Type, Size, and Unit Price columns in order to speed up the process.

<img width="1654" height="50" alt="image" src="https://github.com/user-attachments/assets/d69af84d-9160-4214-a297-211d3256dcc4" />

Next, I populated the Sales columns by multiplying the entries of the Quantity column with those of the Unit Price column. I then used nested IF() statements to create new columns for the names of the coffee types (Arabica, Robusta, etc.) and roast types (Light, Medium, Dark). I then changed the formatting of the Order Date column to dd-Month-yyyy, added the kilogram unit to the Size column, and the dollar symbol ($) to the Unit Price and Sales columns. Lastly, I added a new column called Loyalty Card to indicate which customers are loyalty card members. After removing any duplicates from the orders worksheet, I converted it to a table for ease of use.

I then created a pivot table as a new worksheet from the orders worksheet data. This pivot table includes the Order Date as rows, Coffee Type Name as its columns, and Sales as its values. I generated a 2D line graph of the coffee sales and adjusted the colors to clearly distinguish the different coffee blends. I added 4 slicers to the graph in order to filter its features: Timeline, Size, Roast Type Name, and Loyalty Card. The main idea is that if I adjusted the slicers, the graph would change and I could determine which blend sells the best. For example, in the second attached image of the dashboard I can see that the Arabica and Robusta blends of light and medium roast have higher sales than than the remaining blends even among customers without loyalty cards.

I then created two more pivot tables as separate worksheets from the orders worksheet data: the sales by country and the top 5 customers. In the pivot table for sales by country, I selected Country as the rows and Sales as the values. I then created a 2D bar graph of the data. In the pivot table for the top 5 customers, I selected Customer Name to the rows and Sales to the values. I then created a 2D bar graph of the data. My idea is that knowing which country buys the most coffee could determine which market to target to improve sales, and knowing which customers are busying the most coffee could indicate the success of the loyalty program. In any case, these graphs could have focused on other information, such as the unit price by country.

The main issues I faced in this project was compiling the dashboard with the pivot tables being in different worksheets. The slicers worked on the total sales line graph but did not work on the other two bar graphs. Because to this issue, I had to combine all three pivot tables in one worksheet and create the dashboard below it. I believe that this issue was due to the formatting of Google Sheets worksheets.

<img width="967" height="40" alt="image" src="https://github.com/user-attachments/assets/47bbb8c7-be58-4b7f-948d-a8a5df64ed11" />

# Insights

My main insight is that the U.S.A purchases significantly more coffee than the U.K and Ireland, specifically the Arabica and Excelsa medium roast blends. I also noticed that members of the loyalty program tend to purchase less coffee than nonmembers. This could be due to customers paying with cash, or even being unaware of the perks that come with subscribing to the loyalty program. However, it's important to note that the coffee sales data is from 2019 to 2022 and that sales could be reduced due to the COVID-19 pandemic.

