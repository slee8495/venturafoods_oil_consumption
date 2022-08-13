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

# Forecast ----
forecast <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Oil Consumption/Forecast.xlsx")

forecast[-1,] -> forecast
colnames(forecast) <- forecast[1, ]
forecast[-1, ] -> forecast

forecast %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  dplyr::select(-na, -na_2, -na_4, -forecast_month_year) %>% 
  dplyr::rename(mfg_loc = product_manufacturing_location,
                sku = product_label_sku,
                description = na_3,
                category_no = product_category,
                category = na_5,
                platform_no = product_platform,
                platform = na_6,
                group_no = product_group,
                group = na_7) %>% 
  dplyr::mutate(adjusted_forecast_cases = replace(adjusted_forecast_cases, is.na(adjusted_forecast_cases), 0),
                adjusted_forecast_pounds_lbs = replace(adjusted_forecast_pounds_lbs, is.na(adjusted_forecast_pounds_lbs), 0),
                sku = gsub("-", "", sku),
                ref = paste0(location, "_", sku),
                mfg_ref = paste0(mfg_loc, "_", sku)) %>% 
  dplyr::relocate(ref, mfg_ref) -> forecast



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


# mfg location ----
fg_ref_mfg_ref <- read_excel("S:/Supply Chain Projects/RStudio/BoM/Master formats/FG_On_Hand/FG_ref_to_mfg_ref.xlsx")
fg_ref_mfg_ref %>%
  janitor::clean_names() %>%
  readr::type_convert() %>%
  dplyr::mutate(ref = gsub("-", "_", ref),
                campus_ref = gsub("-", "_", campus_ref),
                mfg_ref = gsub("-", "_", mfg_ref)) %>%
  tidyr::separate(ref, c("1", "2"), sep = "_") %>%
  rename(location = "1") %>%
  dplyr::select(-"2") %>%
  dplyr::mutate(ref = paste0(location, "_", sku)) %>%
  dplyr::relocate(location, mfg_loc, ref, campus_ref, mfg_ref) %>%
  dplyr::mutate(location = as.integer(location)) -> fg_ref_mfg_ref



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





forecast %>% 
  dplyr::left_join(oil_included_sku_2) %>% 
  dplyr::filter(!is.na(oil_included)) %>% 
  dplyr::select(-oil_included) %>% 
  dplyr::mutate(ref = paste0(location, "_", sku)) %>% 
  dplyr::relocate(ref) -> forecast_with_oil



# lbs. only
forecast_with_oil %>% 
  dplyr::select(-adjusted_forecast_cases) -> forecast_with_oil


########################## actual sales & open orders ############################

# Open order cases (make sure with your date range)
# https://edgeanalytics.venturafoods.com/MicroStrategyLibrary/app/DF007F1C11E9B3099BB30080EF7513D2/B226EB613542F97E70A294AB6D55B803/K53--K46

# Sku Actual Shipped (make sure with your date range)
# https://edgeanalytics.venturafoods.com/MicroStrategyLibrary/app/DF007F1C11E9B3099BB30080EF7513D2/BBAA886ACF43D82757EE568F91EEB679/K53--K46

## open orders ----
open_order <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Oil Consumption/Open Orders - 1 Month (3).xlsx")

open_order[-1, ] -> open_order
colnames(open_order) <- open_order[1, ]
open_order[-1, ] -> open_order

open_order %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  dplyr::rename(sku = product_label_sku,
                description =na_2,
                open_order_cases = oo_open_order_cases,
                open_order_net_lbs = oo_net_pounds_lbs) %>% 
  dplyr::mutate(sku = gsub("-", "", sku),
                sales_order_requested_ship_date = as.Date(sales_order_requested_ship_date, origin = "1899-12-30"),
                open_order_cases = replace(open_order_cases, is.na(open_order_cases), 0),
                ref = paste0(location, "_", sku)) %>%
  dplyr::left_join(fg_ref_mfg_ref %>%  dplyr::select(ref, mfg_loc), by = "ref") %>%
  dplyr::mutate(mfg_ref = paste0(mfg_loc, "_", sku)) %>% 
  dplyr::relocate(ref, mfg_ref) %>%
  dplyr::relocate(mfg_loc, .after = location) %>% 
  dplyr::select(-na, -description, -na_3) -> open_order

open_order %>% 
  dplyr::group_by(ref) %>% 
  dplyr::summarise(open_order_net_lbs = sum(open_order_net_lbs)) %>% 
  dplyr::mutate(open_order_net_lbs = replace(open_order_net_lbs, is.na(open_order_net_lbs), 0)) -> open_order_pivot



## sku_actual ----
sku_actual <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Oil Consumption/Sku Actual Shipped (2).xlsx")

sku_actual[-1, ] -> sku_actual
colnames(sku_actual) <- sku_actual[1, ]
sku_actual[-1, ] -> sku_actual

sku_actual %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  dplyr::rename(sku = product_label_sku,
                actual_shipped_cases = cases,
                actual_shipped_lbs = net_pounds_lbs) %>% 
  dplyr::select(sku, location, actual_shipped_lbs) %>% 
  dplyr::mutate(sku = gsub("-", "", sku),
                ref = paste0(location, "_", sku)) %>% 
  dplyr::left_join(fg_ref_mfg_ref %>% dplyr::select(ref, mfg_loc), by = "ref") %>%
  dplyr::mutate(mfg_ref = paste0(mfg_loc, "_", sku)) %>% 
  dplyr::relocate(ref, mfg_ref) %>% 
  dplyr::relocate(mfg_loc, .after = location) -> sku_actual

sku_actual %>% 
  dplyr::group_by(ref) %>% 
  dplyr::summarise(actual_shipped_lbs = sum(actual_shipped_lbs)) %>% 
  dplyr::mutate(actual_shipped_lbs = replace(actual_shipped_lbs, is.na(actual_shipped_lbs), 0)) -> sku_actual_pivot


# combine with dsx_with_oil x open_order

forecast_with_oil %>% 
  dplyr::left_join(open_order_pivot, by = "ref") %>% 
  dplyr::left_join(sku_actual_pivot, by = "ref") %>% 
  dplyr::mutate(open_order_net_lbs = replace(open_order_net_lbs, is.na(open_order_net_lbs), 0),
                actual_shipped_lbs = replace(actual_shipped_lbs, is.na(actual_shipped_lbs), 0),
                open_order_actual_shipped = open_order_net_lbs + actual_shipped_lbs,
                adjusted_forecast_pounds_lbs = round(adjusted_forecast_pounds_lbs, 0)) -> oil_comsumption_comparison


# consumptions
oil_comsumption_comparison %>% 
  dplyr::mutate(consumptions = open_order_actual_shipped / adjusted_forecast_pounds_lbs) %>% 
  dplyr::mutate(consumptions = replace(consumptions, is.na(consumptions), 0),
                consumptions = replace(consumptions, is.nan(consumptions), 0),
                consumptions = replace(consumptions, is.infinite(consumptions), 0),
                consumptions = ifelse(adjusted_forecast_pounds_lbs == 0 & open_order_actual_shipped > 0, "forecasted 0, but sales happened", sprintf("%1.2f%%", 100*consumptions))   ) -> oil_comsumption_comparison


sum(oil_comsumption_comparison$open_order_actual_shipped) / sum(oil_comsumption_comparison$adjusted_forecast_pounds_lbs)


# final touch
oil_comsumption_comparison %>% 
  dplyr::mutate(ref = gsub("_", "-", ref),
                mfg_ref = gsub("_", "-", mfg_ref)) -> oil_comsumption_comparison

str(oil_comsumption_comparison)
# column rename
oil_comsumption_comparison_final <- oil_comsumption_comparison
colnames(oil_comsumption_comparison_final)[1] <- "ref"
colnames(oil_comsumption_comparison_final)[2] <- "mfg_ref"
colnames(oil_comsumption_comparison_final)[3] <- "Location"
colnames(oil_comsumption_comparison_final)[4] <- "mfg_Location"
colnames(oil_comsumption_comparison_final)[5] <- "SKU (FG)"
colnames(oil_comsumption_comparison_final)[6] <- "Description"
colnames(oil_comsumption_comparison_final)[7] <- "Label"
colnames(oil_comsumption_comparison_final)[8] <- "Category"
colnames(oil_comsumption_comparison_final)[9] <- "Category Name"
colnames(oil_comsumption_comparison_final)[10] <- "Platform"
colnames(oil_comsumption_comparison_final)[11] <- "Platform Name"
colnames(oil_comsumption_comparison_final)[12] <- "Group"
colnames(oil_comsumption_comparison_final)[13] <- "Group Name"
colnames(oil_comsumption_comparison_final)[14] <- "Adjusted Forecast Pounds (lbs.)"
colnames(oil_comsumption_comparison_final)[15] <- "Open Order Net Pounds (lbs.)"
colnames(oil_comsumption_comparison_final)[16] <- "Actual Shipped (Previous month)"
colnames(oil_comsumption_comparison_final)[17] <- "Open Order lbs. + Actual Shipped lbs."
colnames(oil_comsumption_comparison_final)[18] <- "Consumptions"

writexl::write_xlsx(oil_comsumption_comparison_final, "oil_compsumtion_comparison.xlsx")




# now to do
# check if open_order and sku_actual has all the lcoation correctly (should I do extra work like Linda did?)



