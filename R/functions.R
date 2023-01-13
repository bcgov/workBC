# a wrapper function for the common pivot_longer applied to the three main inputs.
pl_wrap <- function(tbbl) {
  tbbl %>%
    pivot_longer(
      cols = -c(noc, description, industry, variable, geographic_area),
      names_to = "year",
      values_to = "value"
    )
}
# takes a string, cleans it, and converts to a factor
make_clean_factor <- function(strng) {
  strng %>%
    str_replace_all("\t", "") %>%
    trimws() %>%
    str_to_lower() %>%
    str_replace_all(" ", "_") %>%
    factor()
}
# takes a tibble, converts column names to camel_case, and converts character columns to clean factors.
clean_tbbl <- function(tbbl) {
  tbbl <- tbbl %>%
    mutate(across(where(is.character), make_clean_factor))
  colnames(tbbl) <- if_else(sapply(colnames(tbbl), is.numeric),
    colnames(tbbl),
    str_to_lower(str_replace_all(colnames(tbbl), " ", "_")))
  tbbl
}
# for a given variable, calculates EITHER first five year, second five year and ten year cagrs and returns a tibble OR just the 10 year as a vector.
get_cagrs <- function(tbbl, var, all) {
  tbbl <- tbbl[tbbl$variable == var, ]
  current_value <- pull(tbbl[tbbl$year == current_year, "value"])
  ten_value <- pull(tbbl[tbbl$year == current_plus_10, "value"])
  ten_year <- round((ten_value / current_value)^(.1) - 1, 3) * 100
  if (all == TRUE) {
    five_value <- pull(tbbl[tbbl$year == current_plus_5, "value"])
    first_five <- round((five_value / current_value)^(.2) - 1, 3) * 100
    second_five <- round((ten_value / five_value)^(.2) - 1, 3) * 100
    tibble(
      first_five,
      second_five,
      ten_year
    )
  } else {
    ten_year
  }
}
# for a given variable calculate the ten year sum of the value.  Returns a vector
ten_sum <- function(tbbl, var, round_to=-1) {
  round(sum(pull(tbbl[tbbl$variable == var & tbbl$year > current_year & tbbl$year <= current_plus_10, "value"])), round_to)
}
# for a given variable selects out the current, 5 and 10 year values.
current_5_10 <- function(tbbl, var, round_to=-1) {
  current <- round(pull(tbbl[tbbl$variable == var & tbbl$year == current_year, "value"]), round_to)
  five <- round(pull(tbbl[tbbl$variable == var & tbbl$year == current_plus_5, "value"]), round_to)
  ten <- round(pull(tbbl[tbbl$variable == var & tbbl$year == current_plus_10, "value"]), round_to)
  tibble(
    current,
    five,
    ten
  )
}
# for a given variable get the current value
current_value <- function(tbbl, var) {
  round(pull(tbbl[tbbl$variable == var & tbbl$year == current_year, "value"]),-1)
}

# for a given variable aggregate the value by year.
aggregate_by_year <- function(tbbl, var) {
  tbbl <- tbbl[tbbl$variable == var, ]
  tbbl %>%
    group_by(year) %>%
    summarize(value = round(sum(value),-1)) %>%
    mutate(variable = var)
}
# for a given variable calculate the current, five year and ten year shares of total employment
get_shares <- function(tbbl, var) {
  tbbl <- tbbl[tbbl$variable == var, ]
  current_share <- tbbl[tbbl$year == current_year, "value"] / tot_emp[tot_emp$year == current_year, "value"]
  five_share <- tbbl[tbbl$year == current_plus_5, "value"] / tot_emp[tot_emp$year == current_plus_5, "value"]
  ten_share <- tbbl[tbbl$year == current_plus_10, "value"] / tot_emp[tot_emp$year == current_plus_10, "value"]
  tibble(
    current = round(pull(current_share), 3) * 100,
    five = round(pull(five_share), 3) * 100,
    ten = round(pull(ten_share), 3) * 100,
  )
}

# takes a tibble and converts all factors with camel_case labels to Title Character Strings.
camel_to_title <- function(tbbl) {
  tbbl %>%
    rapply(as.character, classes = "factor", how = "replace") %>%
    tibble() %>%
    mutate(across(where(is.character), make_title))
}
#convert camel_case to Title Case
make_title <- function(strng) {
  strng <- str_to_title(str_replace_all(strng, "_", " "))
}
#convert camel_case to Sentence case
make_sentence <- function(strng) {
  strng <- str_to_sentence(str_replace_all(strng, "_", " "))
}

#takes a tibble and adds a header and a footer.
add_foot_head <- function(tbbl, hdr, num_col){
  assertthat::assert_that(num_col %in% 3:4)
  if(num_col==4){
    header <- c("",{{  hdr  }},"","")
    footer <- c("","","","")
  }else{
    header <- c({{  hdr  }},"","")
    footer <- c("","","")
  }
  rbind(header, tbbl, footer)
}
#takes a tibble and adds a header in the last column... because why not?
add_foot_head2 <- function(tbbl, hdr, num_col){
  assertthat::assert_that(num_col %in% 2:3)
  if(num_col==3){
    header <- c("","",{{  hdr  }})
    footer <- c("","","")
  }else{
    header <- c("", {{  hdr  }})
    footer <- c("","")
  }
  rbind(header, tbbl, footer)
}
#take a tibble and adds a header but no footer
add_header <- function(tbbl, var){
  header <- c(paste("Typical education background:",{{  var  }}),"","","","","","","")
  rbind(header, tbbl)
}
smush_names <- function(tbbl){
  paste(tbbl$trade, collapse="/ ")
}
# Function to quickly export data
write_workbook <- function(data, sheetname, startrow, startcol) {
  writeWorksheet(
    wb,
    data,
    sheetname,
    startRow = startrow,
    startCol = startcol,
    header = FALSE
  )
}
