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

# Bulk Oil List ----
# https://edgeanalytics.venturafoods.com/MicroStrategyLibrary/app/DF007F1C11E9B3099BB30080EF7513D2/A00AF850E84EC6F52CFD9DABD1742F03/K53--K46
bulk_oil_list <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Oil Consumption/Oct.2022 Report/Bulk Oil Table (3).xlsx")

bulk_oil_list %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  dplyr::rename(bulk_oil = bulk_oil_type,
                component = base_product) %>% 
  dplyr::select(-x2, -x4) %>% 
  dplyr::mutate(bulk = ifelse(bulk_oil == "UNKNOW" | bulk_oil == "NONE", "N", "Y")) %>% 
  dplyr::mutate(component = sub("^0+", "", component),
                component = as.numeric(component))-> bulk_oil_list




# Forecast dsx (for lag1, use the first day of the month) ----
# Make sure to put the date correctly few below ----
dsx <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/Demand Planning/Demand Planning Team/BI Forecast Backup/2022/DSX Forecast Backup - 2022.12.01.xlsx")

dsx[-1, ] -> dsx
colnames(dsx) <- dsx[1, ]
dsx[-1, ] -> dsx

dsx %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  data.frame() %>% 
  dplyr::filter(forecast_month_year_code == 202212) %>%    ############################# MAKE SURE TO PUT THE DATE CORRECTLY ####################### ----
dplyr::rename(mfg_loc = product_manufacturing_location_code,
              location = location_no,
              sku = product_label_sku_code,
              sku_description = product_label_sku_name,
              category = product_category_name,
              platform = product_platform_name,
              group_no = product_group_code,
              group = product_group_short_name) %>% 
  dplyr::mutate(sku = gsub("-", "", sku),
                ref = paste0(location, "_", sku),
                mfg_ref = paste0(mfg_loc, "_", sku),
                label = stringr::str_sub(sku, 6, 8)) %>% 
  dplyr::select(mfg_ref, mfg_loc, sku, sku_description, label, category, platform, group_no, group, adjusted_forecast_cases,
                adjusted_forecast_pounds_lbs) %>% 
  dplyr::mutate(adjusted_forecast_pounds_lbs = replace(adjusted_forecast_pounds_lbs, is.na(adjusted_forecast_pounds_lbs), 0),
                adjusted_forecast_cases = replace(adjusted_forecast_cases, is.na(adjusted_forecast_cases), 0)) -> forecast 


forecast %>% 
  dplyr::group_by(mfg_ref, mfg_loc, sku, sku_description, label, category, platform, group_no, group) %>% 
  dplyr::summarise(adjusted_forecast_cases = sum(adjusted_forecast_cases),
                   adjusted_forecast_pounds_lbs = sum(adjusted_forecast_pounds_lbs)) -> forecast


forecast %>% 
  dplyr::filter(mfg_loc != 22) %>% 
  dplyr::filter(mfg_loc != 16) -> forecast


# BoM RM to sku ----
rm_to_sku <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR archive/Raw Material Inventory Health (IQR) - 11.02.22.xlsx", 
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





forecast %>% 
  dplyr::left_join(oil_included_sku_2) %>% 
  dplyr::filter(!is.na(oil_included)) %>% 
  dplyr::select(-oil_included) -> forecast_with_oil




# BoM Report ----
bom <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR archive/Raw Material Inventory Health (IQR) - 11.02.22.xlsx", 
                  sheet = "BoM")

bom[-1:-5, ] -> bom
colnames(bom) <- bom[1, ]
bom[-1, ] -> bom

bom %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  data.frame() %>% 
  dplyr::mutate(ref = gsub("-", "_", ref)) %>% 
  dplyr::select(ref, comp_ref, business_unit, parent_item_number, comp_number_labor_code, quantity_w_scrap) %>% 
  dplyr::rename(mfg_loc = business_unit,
                sku = parent_item_number,
                component = comp_number_labor_code) %>% 
  dplyr::mutate(mfg_ref = paste0(mfg_loc, "_", sku),
                mfg_comp_ref = paste0(mfg_loc, "_", component)) %>% 
  dplyr::left_join(oil_list_3) %>% 
  dplyr::filter(!is.na(oil)) %>% 
  dplyr::select(-oil, -mfg_loc) %>% 
  dplyr::relocate(ref, comp_ref, mfg_ref, mfg_comp_ref, sku, component, quantity_w_scrap) %>% 
  dplyr::mutate(quantity_w_scrap = round(quantity_w_scrap, 2)) -> bom



########################## actual sales & open orders ############################

# Open order cases (make sure with your date range) 
# https://edgeanalytics.venturafoods.com/MicroStrategyLibrary/app/DF007F1C11E9B3099BB30080EF7513D2/B226EB613542F97E70A294AB6D55B803/K53--K46
# oil consumption tab in MS


## open orders ----
open_order <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Oil Consumption/12 month rolling report/Jan.2023/Open Orders - 1 Month (18).xlsx")

open_order[-1, ] -> open_order
colnames(open_order) <- open_order[1, ]
open_order[-1, ] -> open_order

open_order %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  dplyr::rename(location_name = na,
                mfg_loc = product_manufacturing_location,
                mfg_loc_name = na_2,
                component = base_product,
                sku = product_label_sku,
                description = na_3,
                category = na_4,
                category_no = product_category,
                open_order_net_lbs = oo_net_pounds_lbs,
                open_order_cases = oo_cases) %>% 
  dplyr::mutate(sales_order_requested_ship_date = as.Date(sales_order_requested_ship_date, origin = "1899-12-30"),
                open_order_cases = replace(open_order_cases, is.na(open_order_cases), 0),
                sku = gsub("-", "", sku), 
                ref = paste0(location, "_", sku),
                mfg_ref = paste0(mfg_loc, "_", sku)) %>%
  dplyr::relocate(ref, mfg_ref) %>%
  dplyr::relocate(mfg_loc, .after = location) -> open_order

open_order %>% 
  dplyr::group_by(mfg_ref) %>% 
  dplyr::summarise(open_order_net_lbs = sum(open_order_net_lbs),
                   open_order_cases = sum(open_order_cases)) %>% 
  dplyr::mutate(open_order_net_lbs = replace(open_order_net_lbs, is.na(open_order_net_lbs), 0),
                open_order_cases = replace(open_order_cases, is.na(open_order_cases), 0)) -> open_order_pivot




# Sku Actual Shipped (make sure with your date range)
# https://edgeanalytics.venturafoods.com/MicroStrategyLibrary/app/DF007F1C11E9B3099BB30080EF7513D2/7D421DDA4D4411DA73B4469771826BD9/W62--K46

## sku_actual ----
sku_actual <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Oil Consumption/12 month rolling report/Jan.2023/Order and Shipped History (10).xlsx")


sku_actual %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>%
  dplyr::rename(mfg_loc = product_manufacturing_location,
                sku = product_label_sku,
                actual_shipped_cases = cases,
                actual_shipped_lbs = net_pounds_lbs,
                year = calendar_year,
                month = calendar_month_no) %>% 
  dplyr::select(year, month, sku, mfg_loc, actual_shipped_lbs, actual_shipped_cases) %>% 
  dplyr::mutate(sku = gsub("-", "", sku),
                mfg_ref = paste0(mfg_loc, "_", sku)) -> sku_actual

sku_actual %>% 
  dplyr::group_by(year, month, mfg_ref) %>% 
  dplyr::summarise(actual_shipped_lbs = sum(actual_shipped_lbs),
                   actual_shipped_cases = sum(actual_shipped_cases)) %>% 
  dplyr::mutate(actual_shipped_lbs = replace(actual_shipped_lbs, is.na(actual_shipped_lbs), 0),
                actual_shipped_cases = replace(actual_shipped_cases, is.na(actual_shipped_cases), 0)) -> sku_actual_pivot


# combine with dsx_with_oil x open_order

forecast_with_oil %>% 
  dplyr::left_join(open_order_pivot, by = "mfg_ref") %>% 
  dplyr::left_join(sku_actual_pivot, by = "mfg_ref") %>% 
  dplyr::mutate(open_order_cases = replace(open_order_cases, is.na(open_order_cases), 0),
                actual_shipped_cases = replace(actual_shipped_cases, is.na(actual_shipped_cases), 0)) %>% 
  dplyr::mutate(open_order_net_lbs = replace(open_order_net_lbs, is.na(open_order_net_lbs), 0),
                actual_shipped_lbs = replace(actual_shipped_lbs, is.na(actual_shipped_lbs), 0),
                open_order_actual_shipped_lbs = open_order_net_lbs + actual_shipped_lbs,
                open_order_actual_shipped_cases = open_order_cases + actual_shipped_cases,
                adjusted_forecast_pounds_lbs = round(adjusted_forecast_pounds_lbs, 0)) -> oil_comsumption_comparison




################################################# Sales Orders ##################################################
# Input sales orders ----
# https://edgeanalytics.venturafoods.com/MicroStrategyLibrary/app/DF007F1C11E9B3099BB30080EF7513D2/7D421DDA4D4411DA73B4469771826BD9/W62--K46

sales_orders <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Oil Consumption/12 month rolling report/Jan.2023/Order and Shipped History (11).xlsx")

sales_orders %>% 
  janitor::clean_names() %>% 
  dplyr::rename(mfg_loc = product_manufacturing_location,
                sku = product_label_sku,
                description = x6,
                order_qty_final = ordered_final_qty,
                order_qty_original = ordered_original_qty,
                year = calendar_year,
                month = calendar_month_no) %>% 
  dplyr::mutate(order_qty_final = replace(order_qty_final, is.na(order_qty_final), 0),
                order_qty_original = replace(order_qty_original, is.na(order_qty_original), 0),
                sku = gsub("-", "", sku),
                mfg_ref = paste0(mfg_loc, "_", sku)) %>% 
  readr::type_convert() -> sales_orders


sales_orders %>% 
  dplyr::group_by(year, month, mfg_ref) %>% 
  dplyr::summarise(order_qty_final = sum(order_qty_final),
                   order_qty_original = sum(order_qty_original)) %>% 
  dplyr::mutate(order_qty_final = ifelse(order_qty_final < 0, 0, order_qty_final),
                order_qty_original = ifelse(order_qty_original < 0, 0, order_qty_original)) -> sales_orders_pivot



oil_comsumption_comparison %>% 
  dplyr::left_join(sales_orders_pivot, by = "mfg_ref") -> oil_comsumption_comparison


# NA to 0
oil_comsumption_comparison %>% 
  dplyr::mutate(order_qty_final = replace(order_qty_final, is.na(order_qty_final), 0),
                order_qty_original = replace(order_qty_original, is.na(order_qty_original), 0)) -> oil_comsumption_comparison



################################################ second phase #######################################
bom %>% 
  dplyr::select(mfg_ref, component, quantity_w_scrap) -> bom_2



oil_comsumption_comparison %>% 
  dplyr::left_join(bom_2) -> oil_comsumption_comparison_ver2



oil_comsumption_comparison_ver2 %>% 
  dplyr::mutate(forecasted_oil_qty = adjusted_forecast_cases * quantity_w_scrap,
                consumption_qty_actual_shipped = open_order_actual_shipped_cases * quantity_w_scrap,
                consumption_percent_adjusted_actual_shipped = consumption_qty_actual_shipped / forecasted_oil_qty) %>%
  
  dplyr::mutate(consumption_qty_sales_order_qty = order_qty_final * quantity_w_scrap,
                consumption_percent_adjusted_sales_order = consumption_qty_sales_order_qty / forecasted_oil_qty) %>% 
  
  
  dplyr::mutate(consumption_percent_adjusted_actual_shipped = replace(consumption_percent_adjusted_actual_shipped, is.na(consumption_percent_adjusted_actual_shipped) | is.nan(consumption_percent_adjusted_actual_shipped) | is.infinite(consumption_percent_adjusted_actual_shipped), 0)) %>% 
  dplyr::mutate(consumption_percent_adjusted_actual_shipped = sprintf("%1.2f%%", 100*consumption_percent_adjusted_actual_shipped)) %>% 
  dplyr::mutate(consumption_percent_adjusted_sales_order = replace(consumption_percent_adjusted_sales_order, is.na(consumption_percent_adjusted_sales_order) | is.nan(consumption_percent_adjusted_sales_order) | is.infinite(consumption_percent_adjusted_sales_order), 0)) %>% 
  dplyr::mutate(consumption_percent_adjusted_sales_order = sprintf("%1.2f%%", 100*consumption_percent_adjusted_sales_order)) %>% 
  
  
  dplyr::mutate(diff_between_forecast_actual =  forecasted_oil_qty - consumption_qty_actual_shipped,
                diff_between_forecast_original = forecasted_oil_qty - consumption_qty_sales_order_qty) %>% 
  
  
  dplyr::arrange(mfg_ref) %>% 
  dplyr::relocate(component, .after = group) -> oil_comsumption_comparison_ver2




oil_list %>% 
  dplyr::select(component, category) %>% 
  dplyr::rename(oil_description = category) -> oil_desc


oil_comsumption_comparison_ver2 %>% 
  dplyr::left_join(oil_desc) %>% 
  dplyr::relocate(oil_description, .after = component) -> oil_comsumption_comparison_ver2

oil_comsumption_comparison_ver2 %>% 
  dplyr::filter(mfg_loc != "-1") -> oil_comsumption_comparison_ver2


# add bulk_oil column
bulk_oil_list %>% 
  dplyr::select(component, bulk) -> bulk_oil_list_merge

oil_comsumption_comparison_ver2 %>% 
  dplyr::left_join(bulk_oil_list_merge, by = "component") %>% 
  dplyr::mutate(bulk = ifelse(is.na(bulk), "N", bulk)) -> oil_comsumption_comparison_ver2




# Duplicated values delete
oil_comsumption_comparison_ver2[!duplicated(oil_comsumption_comparison_ver2[,c("mfg_ref", "sku", "component", "quantity_w_scrap")]),] -> oil_comsumption_comparison_ver2



#################################################################################################################################################
#################################################################################################################################################

# final round up
oil_comsumption_comparison_ver2 %>% 
  dplyr::mutate(adjusted_forecast_cases = round(adjusted_forecast_cases, 0),
                forecasted_oil_qty = round(forecasted_oil_qty, 0),
                consumption_qty_actual_shipped = round(consumption_qty_actual_shipped, 0),
                diff_between_forecast_actual = round(diff_between_forecast_actual, 0),
                consumption_qty_sales_order_qty = round(consumption_qty_sales_order_qty, 0),
                diff_between_forecast_original = round(diff_between_forecast_original, 0)) -> oil_comsumption_comparison_ver2


# final touch
oil_comsumption_comparison_ver2 %>% 
  dplyr::mutate(mfg_ref = gsub("_", "-", mfg_ref)) -> oil_comsumption_comparison_ver2 

# column rename
oil_comsumption_comparison_final <- oil_comsumption_comparison_ver2

oil_comsumption_comparison_final %>% 
  dplyr::select(mfg_ref, mfg_loc, sku, sku_description, label, category, platform, group_no, group, component, oil_description, bulk,
                quantity_w_scrap, adjusted_forecast_cases, forecasted_oil_qty, 
                open_order_cases, actual_shipped_cases, open_order_actual_shipped_cases, 
                consumption_qty_actual_shipped, consumption_percent_adjusted_actual_shipped,
                diff_between_forecast_actual, order_qty_final, order_qty_original, consumption_qty_sales_order_qty, 
                consumption_percent_adjusted_sales_order, diff_between_forecast_original) -> oil_comsumption_comparison_final


str(oil_comsumption_comparison_final)

colnames(oil_comsumption_comparison_final)[1] <- "mfg ref"
colnames(oil_comsumption_comparison_final)[2] <- "mfg Location"
colnames(oil_comsumption_comparison_final)[3] <- "SKU (FG)"
colnames(oil_comsumption_comparison_final)[4] <- "Description"
colnames(oil_comsumption_comparison_final)[5] <- "Label"
colnames(oil_comsumption_comparison_final)[6] <- "Category"
colnames(oil_comsumption_comparison_final)[7] <- "Platform"
colnames(oil_comsumption_comparison_final)[8] <- "Group Code"
colnames(oil_comsumption_comparison_final)[9] <- "Group Name"
colnames(oil_comsumption_comparison_final)[10] <- "Component (Oil)"
colnames(oil_comsumption_comparison_final)[11] <- "Oil Description"
colnames(oil_comsumption_comparison_final)[12] <- "Bulk?"
colnames(oil_comsumption_comparison_final)[13] <- "Quantity w/Scrap"
colnames(oil_comsumption_comparison_final)[14] <- "Adjusted Forecast Cases"
colnames(oil_comsumption_comparison_final)[15] <- "Forecasted Oil Qty"
colnames(oil_comsumption_comparison_final)[16] <- "Open Order Cases"
colnames(oil_comsumption_comparison_final)[17] <- "Actual Shipped Cases"
colnames(oil_comsumption_comparison_final)[18] <- "Open Order Cases + Actual Shipped Cases"
colnames(oil_comsumption_comparison_final)[19] <- "Consumption Quantity (Open Order + Actual Shipped)"
colnames(oil_comsumption_comparison_final)[20] <- "Consumption % (by Adjusted forecast - Open Order + Actual Shipped)"
colnames(oil_comsumption_comparison_final)[21] <- "Diff (Forecasted - Actual Shipped)"
colnames(oil_comsumption_comparison_final)[22] <- "Sales Order Qty Final (Cases)"
colnames(oil_comsumption_comparison_final)[23] <- "Sales Order Qty Original (Cases)"
colnames(oil_comsumption_comparison_final)[24] <- "Consumption Quantity (Original Sales Order Qty)"
colnames(oil_comsumption_comparison_final)[25] <- "Consumption % (by Adjusted forecast - Original Sales Order Qty)"
colnames(oil_comsumption_comparison_final)[26] <- "Diff (Forecasted - Original Sales Order)"



writexl::write_xlsx(oil_comsumption_comparison_final, "oil_consumption_comparison.xlsx")




# I found the way!!!
# First, I need to create a mega file.. by repeating 12 times. 

