library(readxl)
library(here)
library(janitor)
library(tidytable)
library(purrr)

natamen_path <- here("data", "updated_natamenscale.xlsx")

# Read in various header rows and extract names
(cnames_1 <- read_excel(natamen_path, n_max = 0, range = "A4:AE4") %>%
    names())
(cnames_2 <- read_excel(natamen_path, n_max = 0, range = "A5:AE5") %>%
    names())
(cnames_3 <- read_excel(natamen_path, n_max = 0, range = "A6:AE6") %>%
    names())
(cnames_4 <- read_excel(natamen_path, n_max = 0, range = "A7:AE7") %>%
    names())

cleaned_cnames <- list(cnames_1, cnames_2, cnames_3, cnames_4) %>%
    map(~ gsub("\\.\\.\\.\\d+", "", .))

cnames <- map_chr(1:31, ~ reduce(map(cleaned_cnames, .x), paste, sep = "_")) %>%
    gsub("\\_+", "", .) %>%
    gsub("\\-+", "", .)

natamen_data <- read_excel(natamen_path,
    range = "A8:AE3118",
    col_names = cnames
) %>%
    clean_names() %>%
    remove_empty(which = "cols")
