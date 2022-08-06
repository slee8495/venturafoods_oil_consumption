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


# previous forecast month ----
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
  dplyr::filter(!is.na(oil_included)) -> dsx_with_oil

dsx_with_oil

########################## actual sales & open orders ############################

# Open order cases (make sure with your date range)
# https://edgeanalytics.venturafoods.com/MicroStrategyLibrary/app/DF007F1C11E9B3099BB30080EF7513D2/B226EB613542F97E70A294AB6D55B803/K53--K46

# Sku Actual Shipped (make sure with your date range)
# https://edgeanalytics.venturafoods.com/MicroStrategyLibrary/app/DF007F1C11E9B3099BB30080EF7513D2/BBAA886ACF43D82757EE568F91EEB679/K53--K46

open_order <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Oil Consumption/Open Orders - 1 Month (1).xlsx")

sku_actual <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Oil Consumption/Sku Actual Shipped (1).xlsx")


# now to do
# check if open_order and sku_actual has all the lcoation correctly (should I do extra work like Linda did?)
# connect them, build a query



