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

# Forecast dsx (for lag1, use the first day of the month) ----
dsx <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/Demand Planning/Demand Planning Team/BI Forecast Backup/2022/DSX Forecast Backup - 2022.08.01.xlsx")

dsx[-1, ] -> dsx
colnames(dsx) <- dsx[1, ]
dsx[-1, ] -> dsx

dsx %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  data.frame() %>% 
  dplyr::filter(forecast_month_year_code == 202208) %>%    ############################# MAKE SURE TO PUT THE DATE CORRECTLY ####################### ----
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
  dplyr::select(ref, mfg_ref, location, mfg_loc, sku, sku_description, label, category, platform, group_no, group,
                stat_forecast_pounds_lbs, adjusted_forecast_pounds_lbs) %>% 
  dplyr::mutate(stat_forecast_pounds_lbs = replace(stat_forecast_pounds_lbs, is.na(stat_forecast_pounds_lbs), 0),
                adjusted_forecast_pounds_lbs = replace(adjusted_forecast_pounds_lbs, is.na(adjusted_forecast_pounds_lbs), 0)) -> forecast 





# BoM RM to sku ----
rm_to_sku <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/Raw Material Inventory Health (IQR) - 08.29.22.xlsx", 
                        sheet = "RM to SKU")

rm_to_sku %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  dplyr::select(2:4) %>%
  dplyr::rename(component = comp_number_labor_code,
                comp_description = comp_description_3,
                sku = parent_item_number) %>% 
  dplyr::filter(!is.na(component)) -> rm_to_sku


# # mfg location (This step can be eliminated) ----
# fg_ref_mfg_ref <- read_excel("S:/Supply Chain Projects/RStudio/BoM/Master formats/FG_On_Hand/FG_ref_to_mfg_ref.xlsx")
# fg_ref_mfg_ref %>%
#   janitor::clean_names() %>%
#   readr::type_convert() %>%
#   dplyr::mutate(ref = gsub("-", "_", ref),
#                 campus_ref = gsub("-", "_", campus_ref),
#                 mfg_ref = gsub("-", "_", mfg_ref)) %>%
#   tidyr::separate(ref, c("1", "2"), sep = "_") %>%
#   rename(location = "1") %>%
#   dplyr::select(-"2") %>%
#   dplyr::mutate(ref = paste0(location, "_", sku)) %>%
#   dplyr::relocate(location, mfg_loc, ref, campus_ref, mfg_ref) %>%
#   dplyr::mutate(location = as.integer(location)) -> fg_ref_mfg_ref



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
bom <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/Raw Material Inventory Health (IQR) - 08.29.22.xlsx", 
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
  dplyr::relocate(ref, comp_ref, mfg_ref, mfg_comp_ref, sku, component, quantity_w_scrap) -> bom



########################## actual sales & open orders ############################

# Open order cases (make sure with your date range) 
# https://edgeanalytics.venturafoods.com/MicroStrategyLibrary/app/DF007F1C11E9B3099BB30080EF7513D2/B226EB613542F97E70A294AB6D55B803/K53--K46
# oil consumption tab in MS


## open orders ----
open_order <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Oil Consumption/Open Orders - 1 Month (11).xlsx")

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
  dplyr::summarise(open_order_net_lbs = sum(open_order_net_lbs)) %>% 
  dplyr::mutate(open_order_net_lbs = replace(open_order_net_lbs, is.na(open_order_net_lbs), 0)) -> open_order_pivot


# Sku Actual Shipped (make sure with your date range)
# https://edgeanalytics.venturafoods.com/MicroStrategyLibrary/app/DF007F1C11E9B3099BB30080EF7513D2/BBAA886ACF43D82757EE568F91EEB679/K53--K46

## sku_actual ----
sku_actual <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Oil Consumption/Sku Actual Shipped.xlsx")

sku_actual[-1, ] -> sku_actual
colnames(sku_actual) <- sku_actual[1, ]
sku_actual[-1, ] -> sku_actual

sku_actual %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  dplyr::rename(location_name = na,
                mfg_loc = product_manufacturing_location,
                mfg_loc_name = na_2,
                component = base_product,
                description = na_3,
                sku = product_label_sku,
                category = na_5,
                actual_shipped_cases = cases,
                actual_shipped_lbs = net_pounds_lbs) %>% 
  dplyr::select(sku, location, mfg_loc, actual_shipped_lbs) %>% 
  dplyr::mutate(sku = gsub("-", "", sku),
                ref = paste0(location, "_", sku),
                mfg_ref = paste0(mfg_loc, "_", sku)) -> sku_actual

sku_actual %>% 
  dplyr::group_by(mfg_ref) %>% 
  dplyr::summarise(actual_shipped_lbs = sum(actual_shipped_lbs)) %>% 
  dplyr::mutate(actual_shipped_lbs = replace(actual_shipped_lbs, is.na(actual_shipped_lbs), 0)) -> sku_actual_pivot


# combine with dsx_with_oil x open_order

forecast_with_oil %>% 
  dplyr::left_join(open_order_pivot, by = "mfg_ref") %>% 
  dplyr::left_join(sku_actual_pivot, by = "mfg_ref") %>% 
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



################################################ second phase #######################################
bom %>% 
  dplyr::select(sku, component, quantity_w_scrap) -> bom_2

oil_comsumption_comparison %>% 
  dplyr::left_join(bom_2) %>% 
  dplyr::select(-consumptions) -> oil_comsumption_comparison_ver2


oil_comsumption_comparison_ver2 %>% 
  dplyr::mutate(consumption_qty = (open_order_actual_shipped * quantity_w_scrap),
                consumption_percent_stat = (stat_forecast_pounds_lbs * quantity_w_scrap) / consumption_qty,
                consumption_percent_adjusted = (adjusted_forecast_pounds_lbs * quantity_w_scrap) / consumption_qty,
                diff_in_consumption_stat =  (stat_forecast_pounds_lbs * quantity_w_scrap) - consumption_qty,
                diff_in_consumption_adjusted =  (adjusted_forecast_pounds_lbs * quantity_w_scrap) - consumption_qty) %>% 
  dplyr::mutate(consumption_percent_stat = replace(consumption_percent_stat, is.na(consumption_percent_stat) | is.nan(consumption_percent_stat) | is.infinite(consumption_percent_stat), 0),
                consumption_percent_adjusted = replace(consumption_percent_adjusted, is.na(consumption_percent_adjusted) | is.nan(consumption_percent_adjusted) | is.infinite(consumption_percent_adjusted), 0)) %>% 
  dplyr::mutate(consumption_percent_stat = sprintf("%1.2f%%", 100*consumption_percent_stat),
                consumption_percent_adjusted = sprintf("%1.2f%%", 100*consumption_percent_adjusted)) %>% 
  dplyr::select(-ref) %>% 
  dplyr::arrange(location, mfg_loc) -> oil_comsumption_comparison_ver2





#################################################################################################################################################
#################################################################################################################################################


# final touch
oil_comsumption_comparison_ver2 %>% 
  dplyr::mutate(mfg_ref = gsub("_", "-", mfg_ref)) -> oil_comsumption_comparison_ver2

# column rename
oil_comsumption_comparison_final <- oil_comsumption_comparison_ver2

colnames(oil_comsumption_comparison_final)[1] <- "mfg ref"
colnames(oil_comsumption_comparison_final)[2] <- "Location"
colnames(oil_comsumption_comparison_final)[3] <- "mfg Location"
colnames(oil_comsumption_comparison_final)[4] <- "SKU (FG)"
colnames(oil_comsumption_comparison_final)[5] <- "Description"
colnames(oil_comsumption_comparison_final)[6] <- "Label"
colnames(oil_comsumption_comparison_final)[7] <- "Category"
colnames(oil_comsumption_comparison_final)[8] <- "Platform"
colnames(oil_comsumption_comparison_final)[9] <- "Group Code"
colnames(oil_comsumption_comparison_final)[10] <- "Group Name"
colnames(oil_comsumption_comparison_final)[11] <- "Stat Forecast Pounds (lbs.)"
colnames(oil_comsumption_comparison_final)[12] <- "Adjusted Forecast Pounds (lbs.)"
colnames(oil_comsumption_comparison_final)[13] <- "Open Order Net Pounds (lbs.)"
colnames(oil_comsumption_comparison_final)[14] <- "Actual Shipped (Previous month)"
colnames(oil_comsumption_comparison_final)[15] <- "Open Order lbs. + Actual Shipped lbs."
colnames(oil_comsumption_comparison_final)[16] <- "Component"
colnames(oil_comsumption_comparison_final)[17] <- "Quantity w/Scrap"
colnames(oil_comsumption_comparison_final)[18] <- "Consumption Quantity"
colnames(oil_comsumption_comparison_final)[19] <- "Consumption % (Stat forecast)"
colnames(oil_comsumption_comparison_final)[20] <- "Consumption % (Adjusted forecast)"
colnames(oil_comsumption_comparison_final)[21] <- "Difference in Consumption (Stat forecast)"
colnames(oil_comsumption_comparison_final)[22] <- "Difference in Consumption (Adjusted forecast)"



writexl::write_xlsx(oil_comsumption_comparison_final, "oil_consumption_comparison.xlsx")





