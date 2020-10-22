#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Apr 27 16:27:45 2020

@author: Jesus
"""

import pandas as pd
x1 = pd.ExcelFile("OSSalesData.xlsx")
lines = "="*25 + "\n"
SalesData = x1.parse("Orders")

# =============================================================================
# #1
# #(DO NOT MODIFY OR REMOVE)
# #Professor's original code. We are using it as a reference point.
# =============================================================================

def SubCatProfits():
    SubCatData = SalesData[["Sub-Category", "Profit"]]
    print(lines)
    SubCatProfit = SubCatData.groupby(by = "Sub-Category").sum().sort_values(by = "Profit")
    print(SubCatProfit.head(10))


#2   
# This function will provide the profit for each product name in the 5 lowest performing sub categories , 
#the profit ranges from the least amount of profit to the most 
def ProfitSubCat():
    ProductData = SalesData[["Sub-Category", "Product Name","Profit"]]
    ProfSubCats = ["Tables", "Bookcases", "Supplies", "Fasteners", "Machines"]
    
    for profits in ProfSubCats:
        print(lines)
        ProductInfo = ProductData.loc[ProductData["Sub-Category"]==profits]
        SubCatProfit = ProductInfo.groupby(by = "Product Name").sum().sort_values(by = "Profit")
        print(profits)
        pd.options.display.float_format = '$ {:.2f}'.format
        print(SubCatProfit * 100)    


#3
# This function will provie us with the quantity of products sold in each segment based off of the 4 regions 
def Quantitybyregion():
    Regions = SalesData.Region.unique()
    print(Regions)
    SubCatData = SalesData[["Segment", "Quantity", "Region",]]

    for region in Regions:
        RegionSubCatData = SubCatData.loc[SubCatData["Region"]==region]
        SubCatProfit = RegionSubCatData.groupby(by = "Segment").sum().sort_values(by="Quantity")
        
        print(lines)
        print(region)
        
        print(SubCatProfit.head(10))
# =============================================================================
# #4
# #This function will show the top 10 lowest profiting products in each state that has a buyer in it. 
# #For example if a state only has 2 products, it will only show 2. 
# #If it has 21, it will show the first 10 products and the amount of profit made along with the state's name.
# =============================================================================
        
def StateSubCatProf():
    states = SalesData.State.unique()
    print(states)
    SubCatData = SalesData[["Product Name", "Profit", "State"]]

    for state in states:
        StateSubCatData = SubCatData.loc[SubCatData["State"]==state]
        SubCatProfit = StateSubCatData.groupby(by = "Product Name").sum().sort_values(by="Profit")
        #SubCatDiscount = SubCatData.groupby(by = "Sub-Category").mean().sort_values(by = "Discount")
        print(lines)
        print(state)
        #print(SubCatDiscount)
        print(SubCatProfit.head(10))


# =============================================================================
# #5
# #This function will show the total amount of products bought by each segment. 
# #It will split into the three segments and show the product name and amount bought in the respective segment.        
# =============================================================================

def SegAmt():
    segments = SalesData.Segment.unique()
    print(segments)
    ProdData = SalesData[["Product Name", "Segment", "Quantity"]]
    for segment in segments:
        RegProdData = ProdData.loc[ProdData["Segment"]==segment]
        ProdQuantity = RegProdData.groupby(by = "Product Name").sum().sort_values(by = "Quantity")
        print(lines)
        print(segment)
        print(ProdQuantity)

# =============================================================================
# #6
# This function will show the total profits of the sub-categories from 2016 to 2019
# The output displays the least profiting sub-categories at the top and as slowly descends to the positives
# Most profit will be seen at the bottom
# =============================================================================

def AnnualEfProf():
    SalesDataYear = SalesData
    SalesDataYear["Year"]=SalesDataYear["Order Date"].dt.year
    years = SalesDataYear.Year.unique()
    print(years)
    
    SubCatData=SalesDataYear[["Sub-Category", "Profit", "Year"]]
    
    for year in years:
        SubCatDataByYear = SubCatData.loc[SubCatData["Year"]==year]
        SubCatProfitNoYear = SubCatDataByYear[["Sub-Category", "Profit"]]
        SubCatProfit = SubCatProfitNoYear.groupby(by = "Sub-Category").sum().sort_values(by = "Profit")
        print(lines)
        print(year)
        pd.options.display.float_format = '$ {:.2f}'.format
        print(SubCatProfit.head(10))


# =============================================================================
# #7
# # This function will show the discounts associated with the products in the top 5 underperforming/low profits Sub-Categories. 
# # I derived the Sub-Categories by running the first function which shows the Sub-Categories ascending by profit. 
# # The first 5 show the least profit. 
# # The output will show the Sub-Category name, followed by the product name in each category and the discount associated with that product.
# =============================================================================

def DiscountSubCat():
    ProductData = SalesData[["Sub-Category", "Product Name", "Discount"]]
    DiscSubCats = ["Tables", "Bookcases", "Supplies", "Fasteners", "Machines"]
    
    for discsubcat in DiscSubCats:
        print(lines)
        ProductInfo = ProductData.loc[ProductData["Sub-Category"]== discsubcat]
        SubCatDiscount = ProductInfo.groupby(by = "Product Name").mean().sort_values(by = "Discount")
        print(discsubcat)
        pd.options.display.float_format = '{:.2f}%'.format
        print(SubCatDiscount * 100)
        
        



# =============================================================================
# #8
# This function will show the top 60 customers of the OS shop. 
# After execution, the output will show the customers name, the Sub-Categories they shopped from and the amount of money/sales they spent in each Sub-Category.
# =============================================================================
 
def TopCust():
    SalesDataCust = SalesData
    SalesDataCust["Customer"] = SalesData["Customer Name"]
    customers = SalesDataCust.Customer.unique()
    print(customers)
    SubCatData = SalesData[["Sub-Category","Sales", "Customer"]]
    for cust in customers:
        RegCustData = SubCatData.loc[SubCatData["Customer"]==cust]
        CustSales = RegCustData.groupby(by = "Sub-Category").sum().sort_values(by = "Sales")
        print(lines)
        print(cust)
        print(CustSales.head(60))

# =============================================================================
#  #9
# This function will show the year to year sales from 2016 to 2019 for each category
# Each Sub-category showned is added up together into total sales
# The 4 regions show the same sub-categories, however the discount rate varies from region to region
# =============================================================================

def YearlySubCatSales():
    SalesDataYear = SalesData
    SalesDataYear["Year"] = SalesDataYear["Order Date"].dt.year
    years = SalesDataYear.Year.unique()
    SubCatData = SalesDataYear[["Category", "Sales", "Year"]]
    print(years)
    
    for year in years:
        SubCatDataByYear = SubCatData.loc[SubCatData["Year"]==year]
        SubCatProfitNoYear = SubCatDataByYear[["Category", "Sales"]]
        SubCatProfit = SubCatProfitNoYear.groupby(by = "Category").sum().sort_values(by = "Sales")
        print(lines)
        print(year)
        print(SubCatProfit.head(10))

# =============================================================================
# #10
# This function will display discounts per each region
# The output would allow the viewer to see each sub-category, region, and the discount for the sub-category
# 
# =============================================================================
def DiscountPerRegion():
    regions = SalesData.Region.unique()
    SubCatDiscountRegData = SalesData[["Sub-Category", "Discount", "Region"]]

    for region in regions:
        SubCatInfo = SubCatDiscountRegData.loc[SubCatDiscountRegData["Region"]==region]
        SubCatData = SubCatInfo.groupby(by="Sub-Category").mean().sort_values(by="Discount")
        print(lines)
        print(region)
        print(SubCatData.head(10))

#Main Menu function that displays user commands
def MainMenu():
    print("\n" + "*" * 80)
    print("\n Enter -1- to see Sub-Category Profits" +
          "\n Enter -2- to see the profit of each product in the top 5 underperforming subcategories" +
          "\n Enter -3- to see the quantities of products sold in each region based off of Segments" +
          "\n Enter -4- to see all products by state that are producing the least profits !!!!!!!" +
          "\n Enter -5- to see all underperforming products and quantities bought by each segment !!!!!!" +
          "\n Enter -6- to see all profits of products by year" +
          "\n Enter -7- to see all discounts applied to products of top 5 underperforming subcategories !!!!!!!" +
          "\n Enter -8- to see the details about the top 60 customers for Ofiice Supplies !!!!!!!" +
          "\n Enter -9- to see the Yearly Sales of each Sub-Category" +
          "\n Enter -10- to see the Average Discount for each Region" +
          "\n Enter -11- to exit")
    print("\n" + "*" * 80)
    choice = input("Please enter a number 1 - 11: ") 
    if choice == "1":
        SubCatProfits()
        MainMenu()
    elif choice == "2":
        ProfitSubCat()
        MainMenu()
    elif choice == "3":
        Quantitybyregion()
        MainMenu()
    elif choice == "4":
        StateSubCatProf()
        MainMenu()
    elif choice == "5":
        SegAmt()
        MainMenu()
    elif choice == "6":
        AnnualEfProf()
        MainMenu()
    elif choice == "7":
        DiscountSubCat()
        MainMenu()
    elif choice == "8":
        TopCust()
        MainMenu()
    elif choice == "9":
        YearlySubCatSales()
        MainMenu()
    elif choice == "10":
        DiscountPerRegion()
        MainMenu()
    elif choice == "11":
        exit()
    else:
        print("Invalid input, please try again")
        MainMenu()
print("\nWelcome to Office Solutions Data Analytics System")
MainMenu()