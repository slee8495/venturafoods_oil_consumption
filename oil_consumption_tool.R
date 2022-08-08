library(tidyverse)
library(magrittr)
library(openxlsx)
library(readxl)
library(writexl)
library(reshape2)
library(skimr)
library(janitor)
library(lubridate)


# Oil List ----
oil_list <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Oil Consumption/Oil types AS400 JDE.xlsx")

oil_list %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  dplyr::rename(component = material_number) %>% 
  dplyr::mutate(component = as.double(component)) -> oil_list


# DSX ----
dsx <- read_excel(
  "S:/Global Shared Folders/Large Documents/S&OP/Demand Planning/Demand Planning Team/BI Forecast Backup/DSX Forecast Backup - 2022.07.05.xlsx")

dsx[-1,] -> dsx
colnames(dsx) <- dsx[1, ]
dsx[-1, ] -> dsx

dsx %>% 
  janitor::clean_names() %>% 
  readr::type_convert() -> dsx


# previous forecast month (first line: year-month id) ----
dsx %>% 
  dplyr::filter(forecast_month_year_code == 202207) %>% 
  dplyr::select(location_no, location_name, product_label_sku_code, product_label_sku_name, product_category_name, product_platform_name,
                product_group_code, product_group_short_name, adjusted_forecast_pounds_lbs, adjusted_forecast_cases) %>% 
  dplyr::rename(location = location_no,
                sku = product_label_sku_code,
                description = product_label_sku_name,
                category = product_category_name,
                platform = product_platform_name,
                group = product_group_code,
                group_name = product_group_short_name) %>% 
  dplyr::mutate(adjusted_forecast_pounds_lbs = replace(adjusted_forecast_pounds_lbs, is.na(adjusted_forecast_pounds_lbs), 0),
                adjusted_forecast_cases = replace(adjusted_forecast_cases, is.na(adjusted_forecast_cases), 0)) %>% 
  dplyr::mutate(sku = gsub("-", "", sku)) -> dsx


# BoM RM to sku ----
rm_to_sku <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/Raw Material Inventory Health (IQR) - 08.03.22.xlsx", 
                                                         sheet = "RM to SKU")

rm_to_sku %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  dplyr::select(2:4) %>%
  dplyr::rename(component = comp_number_labor_code,
                comp_description = comp_description_3,
                sku = parent_item_number) %>% 
  dplyr::filter(!is.na(component)) -> rm_to_sku



# combine oil list and RM to Sku
oil_list %>% 
  dplyr::select(component) -> oil_list_2

rm_to_sku %>% 
  dplyr::select(component) -> rm_to_sku_comp

dplyr::intersect(rm_to_sku_comp, oil_list_2) %>% 
  dplyr::mutate(oil = "oil") -> oil_list_3

rm_to_sku %>% 
  dplyr::left_join(oil_list_3, by = "component") %>% 
  filter(!is.na(oil)) %>% 
  dplyr::select(component, sku) -> oil_included_sku



# oil sku extract from dsx sku
oil_included_sku %>% 
  dplyr::select(sku) %>% 
  dplyr::mutate(oil_included = "1") -> oil_included_sku_2

oil_included_sku_2[-which(duplicated(oil_included_sku_2$sku)),] -> oil_included_sku_2

dsx %>% 
  dplyr::left_join(oil_included_sku_2) %>% 
  dplyr::filter(!is.na(oil_included)) %>% 
  dplyr::select(-oil_included) %>% 
  dplyr::mutate(ref = paste0(location, "_", sku)) %>% 
  dplyr::relocate(ref) -> dsx_with_oil

dsx_with_oil

########################## actual sales & open orders ############################

# Open order cases (make sure with your date range)
# https://edgeanalytics.venturafoods.com/MicroStrategyLibrary/app/DF007F1C11E9B3099BB30080EF7513D2/B226EB613542F97E70A294AB6D55B803/K53--K46

# Sku Actual Shipped (make sure with your date range)
# https://edgeanalytics.venturafoods.com/MicroStrategyLibrary/app/DF007F1C11E9B3099BB30080EF7513D2/BBAA886ACF43D82757EE568F91EEB679/K53--K46

## open orders ----
open_order <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Oil Consumption/Open Orders - 1 Month (1).xlsx")

open_order[-1, ] -> open_order
colnames(open_order) <- open_order[1, ]
open_order[-1, ] -> open_order

open_order %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  dplyr::rename(sku = product_label_sku,
                description =na_2) %>% 
  dplyr::mutate(sku = gsub("-", "", sku),
                sales_order_requested_ship_date = as.Date(sales_order_requested_ship_date, origin = "1899-12-30"),
                open_order_cases = replace(open_order_cases, is.na(open_order_cases), 0),
                ref = paste0(location, "_", sku)) %>% 
  dplyr::relocate(ref) %>% 
  dplyr::select(-na, -description, -na_3) -> open_order

open_order %>% 
  dplyr::group_by(ref) %>% 
  dplyr::summarise(open_order_cases = sum(open_order_cases)) %>% 
  dplyr::mutate(open_order_cases = replace(open_order_cases, is.na(open_order_cases), 0)) -> open_order



## sku_actual ----
sku_actual <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Oil Consumption/Sku Actual Shipped (1).xlsx")

sku_actual[-1, ] -> sku_actual
colnames(sku_actual) <- sku_actual[1, ]
sku_actual[-1, ] -> sku_actual

sku_actual %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  dplyr::rename(sku = product_label_sku,
                actual_shipped = cases) %>% 
  dplyr::select(sku, location, actual_shipped) %>% 
  dplyr::mutate(sku = gsub("-", "", sku),
                ref = paste0(location, "_", sku)) %>% 
  dplyr::relocate(ref) -> sku_actual

sku_actual %>% 
  dplyr::group_by(ref) %>% 
  dplyr::summarise(actual_shipped = sum(actual_shipped)) %>% 
  dplyr::mutate(actual_shipped = replace(actual_shipped, is.na(actual_shipped), 0)) -> sku_actual


# combine with dsx_with_oil x open_order

dsx_with_oil %>% 
  dplyr::left_join(open_order, by = "ref") %>% 
  dplyr::left_join(sku_actual, by = "ref") %>% 
  dplyr::mutate(open_order_cases = replace(open_order_cases, is.na(open_order_cases), 0),
                actual_shipped = replace(actual_shipped, is.na(actual_shipped), 0),
                open_order_actual_shipped = open_order_cases + actual_shipped,
                adjusted_forecast_pounds_lbs = round(adjusted_forecast_pounds_lbs, 0),
                ref = gsub("_", "-", ref)) -> oil_comsumption_comparison


sum(oil_comsumption_comparison$open_order_actual_shipped) / sum(oil_comsumption_comparison$adjusted_forecast_cases)




# column rename
oil_comsumption_comparison_final <- oil_comsumption_comparison
colnames(oil_comsumption_comparison_final)[1] <- "ref"
colnames(oil_comsumption_comparison_final)[2] <- "Location"
colnames(oil_comsumption_comparison_final)[3] <- "Location Name"
colnames(oil_comsumption_comparison_final)[4] <- "SKU (FG)"
colnames(oil_comsumption_comparison_final)[5] <- "Description"
colnames(oil_comsumption_comparison_final)[6] <- "Category"
colnames(oil_comsumption_comparison_final)[7] <- "Platform"
colnames(oil_comsumption_comparison_final)[8] <- "Group"
colnames(oil_comsumption_comparison_final)[9] <- "Group Name"
colnames(oil_comsumption_comparison_final)[10] <- "Adjusted Forecast Pounds (lbs.)"
colnames(oil_comsumption_comparison_final)[11] <- "Adjusted Forecast Cases"
colnames(oil_comsumption_comparison_final)[12] <- "Open Order Cases (Previous month)"
colnames(oil_comsumption_comparison_final)[13] <- "Actual Shipped (Previous month"
colnames(oil_comsumption_comparison_final)[14] <- "Open Order Cases + Actual Shipped"


writexl::write_xlsx(oil_comsumption_comparison_final, "oil_compsumtion_comparison.xlsx")




# now to do
# check if open_order and sku_actual has all the lcoation correctly (should I do extra work like Linda did?)


