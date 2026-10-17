## @page { size: A4; margin: 2cm 2cm 2cm 3cm; }
## body { font-family: "Times New Roman", Times, serif; font-size: 14pt; text-align: justify;
##        padding: 2cm 2cm 2cm 3cm; margin: 0 auto; max-width: 1100px; color: #111; }
## p { text-indent: 1.25cm; line-height: 1.5; }
## h1, h2, h3 { text-indent: 0; text-align: left; font-size: 16pt; break-after: avoid; }
## h2, h3 { font-size: 14pt; }
## table { width: 100%; border-collapse: collapse; margin: 12pt 0 18pt; font-size: 14pt; }
## th, td { border-bottom: 1px solid #999; padding: 5pt; text-align: left; overflow-wrap: anywhere; }
## thead { display: table-header-group; }
## img { max-width: 100%; height: auto; }
## .title, .subtitle, div[custom-style="Title"], div[custom-style="Subtitle"] { text-align: center; }
## div[custom-style="Title"] p, div[custom-style="Subtitle"] p { text-align: center; text-indent: 0; }
## .page-break { break-after: page; }
## .caption { text-align: center; text-indent: 0; }
## @media print { body { padding: 0; max-width: none; } #TOC { break-after: page; } }

# Путь к исходному файлу меняется только здесь. Рабочая папка — папка Rmd/R.
xlsx_path <- "data/variant95.xlsx"
out_dir <- "results"
dir.create(out_dir, showWarnings = FALSE, recursive = TRUE)
dir.create(file.path(out_dir, "figures"), showWarnings = FALSE, recursive = TRUE)
options(scipen = 20, knitr.kable.NA = "нет данных")
knitr::opts_chunk$set(echo = FALSE, message = FALSE, warning = FALSE,
                      fig.width = 8.5, fig.height = 5.2, dpi = 180,
                      dev = "png", dev.args = list(type = "cairo"))
library(openxlsx)
library(ggplot2)
library(knitr)
stopifnot(file.exists(xlsx_path))
Data_OG <- openxlsx::read.xlsx(xlsx_path, sheet = "Анкеты", check.names = FALSE)
Data <- Data_OG
names(Data) <- gsub(".", " ", names(Data), fixed = TRUE)
N <- nrow(Data)
for (nm in names(Data)) {
  if (is.character(Data[[nm]])) {
    Data[[nm]] <- trimws(Data[[nm]])
    Data[[nm]][Data[[nm]] == ""] <- NA_character_
  }
}
attrs <- c("Пол", "Возраст", "Образование", "Род занятий",
           "Тип предприятия работодателя", "Тип расселения",
           "Категория дохода на одного члена семьи",
           "Самооценка материального положения семьи")
expected <- c("ID", attrs, paste0("Q", 1:13))
missing_cols <- setdiff(expected, names(Data))
stopifnot(all(c("ID", paste0("Q", 1:13)) %in% names(Data)))
intervals <- c("на 2% и менее", "на 3–5%", "на 6–8%", "на 9–12%",
  "на 13–16%", "на 17–20%", "на 21–25%", "на 26–30%",
  "на 31–40%", "на 41–50%", "на 51% и более")
income_levels <- c("10 000 руб. и менее", "10 001 – 15 000 руб.",
 "15 001 – 20 000 руб.", "20 001 – 25 000 руб.", "25 001 – 30 000 руб.",
 "30 001 – 40 000 руб.", "40 001 – 50 000 руб.", "50 001 – 60 000 руб.",
 "60 001 – 70 000 руб.", "70 001 – 80 000 руб.", "80 001 – 90 000 руб.",
 "90 001 – 100 000 руб.", "100 001 – 120 000 руб.",
 "120 001 – 150 000 руб.", "свыше 150 000 руб.",
 "не знаю, затрудняюсь ответить", "отказ от ответа")
valid <- list(
 "Пол" = c("женщины", "мужчины"),
 "Возраст" = c("18–30 лет", "31–45 лет", "46–60 лет", "старше 60 лет"),
 "Образование" = c("высшее", "ниже среднего", "среднее общее", "среднее специальное"),
 "Род занятий" = c("не работают и не ищут работу", "не работают, но ищут работу",
                   "неработающие пенсионеры", "работающие по найму", "самозанятые"),
 "Тип предприятия работодателя" = c("государственное, с госучастием", "частное, НКО"),
 "Тип расселения" = c("Городское", "Сельское"),
 "Самооценка материального положения семьи" = c("денег не хватает на питание",
  "на питание хватает, на одежду – нет", "на одежду хватает, на крупную бытовую технику – нет",
  "на бытовую технику хватает, на автомобиль – нет", "на автомобиль хватает"),
 Q1 = c("выросли очень сильно", "выросли умеренно", "выросли незначительно",
        "не изменились", "снизились", "затрудняюсь ответить"),
 Q2 = c("вырастут очень сильно", "вырастут умеренно", "вырастут незначительно",
        "не изменятся", "снизятся", "затрудняюсь ответить"),
 Q3 = c("росли быстрее, чем раньше", "росли так же, как и раньше",
  "росли медленнее, чем раньше", "остались на прежнем уровне / не изменились",
  "снижались", "затрудняюсь ответить"),
 Q4 = c(intervals, "затрудняюсь ответить"),
 Q5 = c("будут расти быстрее, чем сейчас", "будут расти так же, как и сейчас",
  "будут расти медленнее, чем сейчас", "останутся на нынешнем уровне / не изменятся",
  "будут снижаться", "затрудняюсь ответить"),
 Q6 = c(intervals, "затрудняюсь ответить"),
 Q7 = c("улучшилось", "осталось без изменений", "ухудшилось", "затрудняюсь ответить"),
 Q8 = c("улучшится", "останется без изменений", "ухудшится", "затрудняюсь ответить"),
 Q9 = c("хорошим", "хорошим, но не во всём", "не хорошим, но и не плохим",
        "плохим, но не во всём", "плохим", "затрудняюсь ответить"),
 Q10 = c("хорошим", "не хорошим, но и не плохим", "плохим", "затрудняюсь ответить"),
 Q11 = c("хорошее", "не хорошее, но и не плохое", "плохое", "затрудняюсь ответить"),
 Q13 = income_levels)
# Условие задавания определяется только допустимым родительским ответом.
eligible <- function(parent, lev) {
  ifelse(is.na(parent) | !(parent %in% lev), NA, parent %in% lev[1:3])
}
asked4 <- eligible(Data$Q3, valid$Q3)
asked6 <- eligible(Data$Q5, valid$Q5)
hired <- ifelse(is.na(Data[["Род занятий"]]) |
 !(Data[["Род занятий"]] %in% valid[["Род занятий"]]), NA,
 Data[["Род занятий"]] == "работающие по найму")
scope <- list(Q4 = asked4, Q6 = asked6, "Тип предприятия работодателя" = hired)
quality <- data.frame(Проверка = character(), Число = integer(), check.names = FALSE)
add_problem <- function(label, n) {
  quality <<- rbind(quality, data.frame(Проверка = label, Число = n, check.names = FALSE))
}
add_problem("Отсутствующие столбцы (единица — столбцы)", length(missing_cols))
add_problem("Отсутствующая характеристика дохода (затронуто анкет)",
            ifelse(attrs[7] %in% missing_cols, N, 0))
add_problem("Пустой ID", sum(is.na(Data$ID)))
add_problem("Строки с повторяющимся непустым ID", sum(!is.na(Data$ID) &
 (duplicated(Data$ID) | duplicated(Data$ID, fromLast = TRUE))))
bad_matrix <- matrix(FALSE, N, length(valid))
for (j in seq_along(valid)) {
 nm <- names(valid)[j]
 if (!(nm %in% names(Data))) next
 bad <- !is.na(Data[[nm]]) & !(Data[[nm]] %in% valid[[nm]])
 bad_matrix[, j] <- bad
 Data[[nm]][bad] <- NA
}
add_problem("Недопустимая категория хотя бы в одном поле (строки)", sum(apply(bad_matrix, 1, any)))
q12 <- suppressWarnings(as.numeric(Data$Q12))
bad12 <- is.na(q12) | !is.finite(q12) | q12 <= 0 | q12 != floor(q12)
add_problem("Q12: пропуск, нечисловое, нецелое или неположительное", sum(bad12))
q12[bad12] <- NA_real_; Data$Q12 <- q12
for (nm in names(scope)) {
 a <- scope[[nm]]; x <- Data[[nm]]
 add_problem(paste(nm, "пусто, хотя поле требуется"), sum(a %in% TRUE & is.na(x)))
 add_problem(paste(nm, "заполнено, хотя поле не задавалось"), sum(a %in% FALSE & !is.na(x)))
}
# Частоты и аудит отдельно; technical не включает недопустимые непустые значения.
freq <- list(); freq_all <- list(); audit <- data.frame()
for (nm in names(valid)) {
 if (!(nm %in% names(Data))) next
 x <- Data[[nm]]; raw <- Data_OG[[nm]]
 a <- if (nm %in% names(scope)) scope[[nm]] else rep(TRUE, N)
 ok <- a %in% TRUE
 f <- as.integer(table(factor(x[ok], levels = valid[[nm]])))
 base <- if (nm %in% c("Q4", "Q6")) sum(ok) else sum(f)
 freq[[nm]] <- data.frame(Категория = valid[[nm]], n = f,
   `Доля, %` = if (base > 0) 100 * f / base else NA_real_, check.names = FALSE)
 missing_raw <- is.na(raw) | trimws(as.character(raw)) == ""
 invalid <- !missing_raw & !(trimws(as.character(raw)) %in% valid[[nm]])
 audit <- rbind(audit, data.frame(Поле = nm, Основание = base,
   Пригодно = sum(f), `Тех. пропуск` = sum(ok & missing_raw),
   `Недопустимо` = sum(ok & invalid), `Структ. пропуск` = sum(a %in% FALSE & missing_raw),
   `Логич. конфликт` = sum(a %in% FALSE & !missing_raw),
   `Условие неизвестно` = sum(is.na(a)), check.names = FALSE))
 if (nm %in% c("Q4", "Q6")) {
   f_all <- c(f, sum(ok & missing_raw), sum(ok & invalid),
     sum(a %in% FALSE), sum(is.na(a)))
   freq_all[[nm]] <- data.frame(Категория = c(valid[[nm]], "Технический пропуск",
    "Недопустимый ответ", "Вопрос не задавался", "Условие задавания неизвестно"),
    n = f_all, `Доля от всех, %` = 100 * f_all / N, check.names = FALSE)
 }
 # При конфликте логики дочернее поле не участвует в расчетах.
 if (nm %in% names(scope)) Data[[nm]][!(a %in% TRUE)] <- NA
}
code_values <- c(-5, 0, 1, 4, 7, 10.5, 14.5, 18.5, 23, 28, 35.5, 45.5, 75)
code_labels <- c("Снижение (условный код)", "Без изменения", intervals)
lower <- c(NA, NA, NA, 2.5, 5.5, 8.5, 12.5, 16.5, 20.5, 25.5, 30.5, 40.5, NA)
width <- c(NA, NA, NA, 3, 3, 4, 4, 4, 5, 5, 10, 10, NA)
recode_infl <- function(parent, child, lev) {
 ans <- rep(NA_real_, length(parent))
 ans[parent %in% lev[4]] <- 0
 ans[parent %in% lev[5]] <- -5
 k <- parent %in% lev[1:3]
 ans[k] <- code_values[3:13][match(child[k], intervals)]
 ans
}
Data$infl_obs <- recode_infl(Data$Q3, Data$Q4, valid$Q3)
Data$infl_exp <- recode_infl(Data$Q5, Data$Q6, valid$Q5)
grouped_median <- function(x, values, L, h, labels = as.character(values)) {
 f <- as.integer(table(factor(x, levels = values))); F <- cumsum(f); n <- sum(f)
 if (n == 0) stop("Нет содержательных наблюдений для медианы")
 k <- which(F >= n / 2)[1]; prev <- ifelse(k == 1, 0, F[k - 1])
 defined <- !is.na(L[k]) && !is.na(h[k])
 med <- if (defined) L[k] + h[k] * (n / 2 - prev) / f[k] else values[k]
 list(N = n, k = k, L = L[k], h = h[k], f = f[k], prev = prev,
  value = med, defined = defined, label = labels[k],
  table = data.frame(Категория = labels, Код = values, n = f, F = F, check.names = FALSE))
}
gobs <- grouped_median(Data$infl_obs, code_values, lower, width, code_labels)
gexp <- grouped_median(Data$infl_exp, code_values, lower, width, code_labels)
infl_summary <- function(x, gm, label) {
 data.frame(Показатель = label, N = sum(!is.na(x)),
 Q1 = unname(quantile(x, .25, na.rm = TRUE, type = 7)),
 Медиана = median(x, na.rm = TRUE),
 Q3 = unname(quantile(x, .75, na.rm = TRUE, type = 7)),
 Уточненная = gm$value, check.names = FALSE)
}
infl_stats <- rbind(infl_summary(Data$infl_obs, gobs, "Наблюдаемая"),
                    infl_summary(Data$infl_exp, gexp, "Ожидаемая"))
income_codes <- c(7000, 12500, 17500, 22500, 27500, 35000, 45000, 55000,
                  65000, 75000, 85000, 95000, 110000, 135000, 200000, NA, NA)
Data$income_hh <- income_codes[match(Data$Q13, income_levels)]
Data$income_pc <- Data$income_hh / Data$Q12
yy <- Data$income_pc[!is.na(Data$income_pc)]
stopifnot(length(yy) > 0)
income_breaks <- seq(0, (floor(max(yy) / 5000) + 1) * 5000, 5000)
income_bin <- cut(yy, breaks = income_breaks, right = FALSE, include.lowest = TRUE)
income_f <- as.integer(table(income_bin)); income_F <- cumsum(income_f)
income_k <- which(income_F >= length(yy) / 2)[1]
income_prev <- ifelse(income_k == 1, 0, income_F[income_k - 1])
income_L <- income_breaks[income_k]
income_gmed <- income_L + 5000 * (length(yy) / 2 - income_prev) / income_f[income_k]
income_series <- data.frame(Интервал = levels(income_bin),
 n = income_f, F = income_F, check.names = FALSE)
income_stats <- data.frame(N = length(yy), Среднее = mean(yy), Медиана = median(yy),
 Q1 = unname(quantile(yy, .25)), Q3 = unname(quantile(yy, .75)),
 IQR = IQR(yy), Уточненная = income_gmed, check.names = FALSE)
family_stats <- data.frame(N = sum(!is.na(Data$Q12)),
 Минимум = min(Data$Q12, na.rm = TRUE), Максимум = max(Data$Q12, na.rm = TRUE),
 Среднее = mean(Data$Q12, na.rm = TRUE), Медиана = median(Data$Q12, na.rm = TRUE))
family_f <- as.data.frame(table(Data$Q12)); names(family_f) <- c("Членов семьи", "n")
family_f$`Доля, %` <- 100 * family_f$n / sum(family_f$n)
positive <- list("улучшилось", "улучшится", c("хорошим", "хорошим, но не во всём"),
                 "хорошим", "хорошее")
negative <- list("ухудшилось", "ухудшится", c("плохим", "плохим, но не во всём"),
                 "плохим", "плохое")
index_names <- c("Материальное положение семьи за прошедший год",
 "Ожидаемое материальное положение семьи через год",
 "Перспективы экономики страны на ближайший год",
 "Перспективы экономики страны на пять лет", "Благоприятность крупных покупок")
calc_indices <- function(d) {
 ans <- data.frame()
 for (j in 1:5) {
  q <- paste0("Q", j + 6); x <- d[[q]]; n <- sum(x %in% valid[[q]])
  p <- sum(x %in% positive[[j]]); m <- sum(x %in% negative[[j]])
  P <- ifelse(n > 0, 100 * p / n, NA); M <- ifelse(n > 0, 100 * m / n, NA)
  ans <- rbind(ans, data.frame(Индекс = paste0("I", j), Вопрос = q,
   n_pos = p, n_neg = m, N = n, P = P, M = M, Значение = 100 + P - M,
   check.names = FALSE))
 }
 ans
}
private <- calc_indices(Data); iv <- private$Значение
ITS <- (iv[1] + iv[5]) / 2; IPO <- (iv[2] + iv[3] + iv[4]) / 3
IPN <- sum(iv) / 5; IPN_check <- (2 * ITS + 3 * IPO) / 5
stopifnot(abs(IPN - IPN_check) < 1e-10, all(iv >= 0 & iv <= 200))
composite <- data.frame(Индекс = c("ИТС", "ИПО", "ИПН", "ИПН: проверочная формула"),
 Значение = c(ITS, IPO, IPN, IPN_check), check.names = FALSE)
group_cuts <- quantile(yy, probs = seq(0, 1, .2), type = 7)
stopifnot(length(unique(group_cuts)) == 6)
group_labels <- paste0("Г", 1:5)
Data$income_group5 <- cut(Data$income_pc, breaks = group_cuts,
 include.lowest = TRUE, right = TRUE, labels = group_labels)
groups <- data.frame(); group_q5 <- data.frame()
min_group <- 30 # Учебный ориентир, не результат проверки значимости.
for (g in group_labels) {
 d <- Data[which(Data$income_group5 == g), ]; j <- match(g, group_labels)
 f <- as.integer(table(factor(d$Q5, levels = valid$Q5))); n <- sum(f)
 group_q5 <- rbind(group_q5, data.frame(Группа = g, Ответ = valid$Q5, n = f,
  `N Q5` = n, `Доля, %` = 100 * f / n, check.names = FALSE))
 ng <- sum(!is.na(d$infl_exp))
 groups <- rbind(groups, data.frame(Группа = g, Нижняя = group_cuts[j],
 Верхняя = group_cuts[j+1], `N группы` = nrow(d), `N Q5` = n,
 `N infl_exp` = ng, Медиана = ifelse(ng > 0, median(d$infl_exp, na.rm = TRUE), NA),
 Пометка = ifelse(ng < min_group, "Малочисленная", "N ≥ 30"), check.names = FALSE))
}
sex_ipn <- data.frame(); sex_private <- data.frame()
for (sx in c("мужчины", "женщины")) {
 d <- Data[which(Data$Пол == sx), ]; ii <- calc_indices(d)
 sex_private <- rbind(sex_private, data.frame(Пол = sx, ii, check.names = FALSE))
 sex_ipn <- rbind(sex_ipn, data.frame(Пол = sx, N = nrow(d),
 `N Q7` = ii$N[1], `N Q8` = ii$N[2], `N Q9` = ii$N[3],
 `N Q10` = ii$N[4], `N Q11` = ii$N[5], ИПН = mean(ii$Значение),
 Пометка = ifelse(min(ii$N) < min_group, "Малочисленная", "N ≥ 30"), check.names = FALSE))
}
usable <- data.frame(Показатель = c(audit$Поле, "Q12", "infl_obs", "infl_exp", "Доход на человека"),
 N = c(audit$Пригодно, sum(!is.na(Data$Q12)), gobs$N, gexp$N, length(yy)), check.names = FALSE)
# Официальные числа из примера друга намеренно НЕ используются: происхождение не подтверждено.
# ЗАПОЛНИТЬ: по точным таблицам Банка России, январь—август 2026 (8 значений).
# Источник: https://www.cbr.ru/Collection/Collection/File/62282/inFOM_26-08.pdf
# Точность и дата получения также заполняются после проверки источника.
official_obs <- rep(NA_real_, 8) # ЗАПОЛНИТЬ: медианы наблюдаемой инфляции, % годовых.
official_exp <- rep(NA_real_, 8) # ЗАПОЛНИТЬ: медианы ожидаемой инфляции, % годовых.
official_ITS <- rep(NA_real_, 8) # ЗАПОЛНИТЬ: индекс текущего состояния, пункты.
official_IPO <- rep(NA_real_, 8) # ЗАПОЛНИТЬ: индекс потребительских ожиданий, пункты.
official_IPN <- rep(NA_real_, 8) # ЗАПОЛНИТЬ: индекс потребительских настроений, пункты.
official_date <- NA_character_ # ЗАПОЛНИТЬ: дата успешного получения, YYYY-MM-DD.
official_precision <- NA_character_ # ЗАПОЛНИТЬ: точность публикации.
months <- c("янв", "фев", "мар", "апр", "май", "июн", "июл", "авг", "сен")
history <- data.frame(Период = paste0("2026-", sprintf("%02d", 1:9)),
 obs = c(official_obs, gobs$value), exp = c(official_exp, gexp$value),
 ITS = c(official_ITS, ITS), IPO = c(official_IPO, IPO), IPN = c(official_IPN, IPN),
 Источник = c(rep("Опубликованный ряд: значение не получено", 8), "Синтетическая выборка"),
 check.names = FALSE)
hist_ok <- !is.na(history$obs) | !is.na(history$exp) | !is.na(history$ITS) |
 !is.na(history$IPO) | !is.na(history$IPN)
history$Источник[1:8][hist_ok[1:8]] <- "Опубликованный ряд (частично, см. NA)"
history$Источник[1:8][complete.cases(history[1:8, 2:6])] <- "Опубликованный ряд"
history$`Дата получения` <- c(rep(official_date, 8), "2026-10-17")
history$Точность <- c(rep(official_precision, 8), "Расчеты без промежуточного округления")
delta <- data.frame(Показатель = c("Наблюдаемая медиана", "Ожидаемая медиана", "ИТС", "ИПО", "ИПН"),
 Август = c(official_obs[8], official_exp[8], official_ITS[8], official_IPO[8], official_IPN[8]),
 Сентябрь = c(gobs$value, gexp$value, ITS, IPO, IPN), check.names = FALSE)
delta$Разность <- delta$Сентябрь - delta$Август
delta$Единица <- c("п. п.", "п. п.", rep("пункты", 3))
delta$Статус <- ifelse(is.na(delta$Август), "Официальное значение не получено", "Рассчитано")
theme_set(theme_minimal(base_size = 12, base_family = "serif") +
 theme(legend.position = "bottom", plot.title.position = "plot",
       panel.grid.minor = element_blank()))
short_month <- c("Очень сильно", "Умеренно", "Незначительно", "Без изменения", "Снижение", "Затрудняюсь")
short_year <- c("Быстрее", "Так же", "Медленнее", "Без изменения", "Снижение", "Затрудняюсь")
pair_data <- function(qs, labels) {
 a <- data.frame()
 for (q in qs) a <- rbind(a, data.frame(Вопрос = q,
  Ответ = factor(labels, levels = labels), Доля = freq[[q]][[3]]))
 a
}
plot_pair <- function(qs, labels) {
 ggplot(pair_data(qs, labels), aes(Ответ, Доля)) + geom_col(fill = "#446a88", width = .72) +
 geom_text(aes(label = sprintf("%.1f", Доля)), vjust = -.25, size = 3) +
 facet_wrap(~Вопрос, nrow = 1) + scale_x_discrete(drop = FALSE) +
 scale_y_continuous(expand = expansion(mult = c(0, .12))) +
 labs(x = "Категория ответа (порядок анкеты)", y = "Доля допустимых ответов, %") +
 theme(axis.text.x = element_text(angle = 38, hjust = 1))
}
fig1 <- plot_pair(c("Q1", "Q2"), short_month)
fig2 <- plot_pair(c("Q3", "Q5"), short_year)
plot_infl <- rbind(data.frame(Переменная = "Наблюдаемая", Код = code_values, n = gobs$table$n),
 data.frame(Переменная = "Ожидаемая", Код = code_values, n = gexp$table$n))
med_lines <- data.frame(Переменная = rep(c("Наблюдаемая", "Ожидаемая"), each = 2),
 Тип = rep(c("Обычная", "Уточненная"), 2),
 x = c(median(Data$infl_obs, na.rm = TRUE), gobs$value,
       median(Data$infl_exp, na.rm = TRUE), gexp$value))
fig3 <- ggplot(plot_infl, aes(Код, n)) + geom_col(fill = "#718da3", width = .65, position = "identity") +
 geom_vline(data = med_lines, aes(xintercept = x, linetype = Тип, color = Тип), linewidth = .8) +
 facet_wrap(~Переменная, ncol = 1) + scale_x_continuous(breaks = code_values, guide = guide_axis(n.dodge = 2)) +
 scale_color_manual(values = c("Обычная" = "#242424", "Уточненная" = "#bc543d")) +
 labs(x = "Кодовая оценка, % годовых (13 значений)", y = "Число ответов, человек",
      color = "Медиана", linetype = "Медиана") +
 theme(axis.text.x = element_text(angle = 45, hjust = 1))
fig4 <- ggplot(data.frame(y = yy), aes(y)) +
 geom_histogram(breaks = income_breaks, closed = "left", fill = "#446a88", color = "white") +
 geom_vline(aes(xintercept = median(yy), linetype = "Обычная медиана"), color = "#b74030", linewidth = .8) +
 labs(x = "Доход, руб. на человека в месяц", y = "Число респондентов, человек", linetype = "")
fig5 <- ggplot(data.frame(Название = factor(paste0("I", 1:5, ": ", index_names),
 levels = rev(paste0("I", 1:5, ": ", index_names))), Значение = iv), aes(Название, Значение)) +
 geom_col(fill = "#446a88", width = .7) + geom_hline(yintercept = 100, linetype = "dashed") +
 geom_text(aes(label = sprintf("%.1f", Значение)), hjust = -.15, size = 3.5) +
 coord_flip() + scale_y_continuous(limits = c(0, max(115, iv + 10))) +
 labs(x = "Частный индекс", y = "Значение, пункты (штриховая линия — 100)") +
 theme(axis.text.y = element_text(size = 10))
fig6 <- ggplot(data.frame(Индекс = factor(c("ИТС", "ИПО", "ИПН"), levels = c("ИТС", "ИПО", "ИПН")),
 Значение = c(ITS, IPO, IPN)), aes(Индекс, Значение)) +
 geom_col(fill = "#446a88", width = .6) + geom_hline(yintercept = 100, linetype = "dashed") +
 geom_text(aes(label = sprintf("%.1f", Значение)), vjust = -.4) +
 scale_y_continuous(limits = c(0, max(115, c(ITS, IPO, IPN) + 10))) +
 labs(x = "ИТС — текущее состояние; ИПО — ожидания; ИПН — настроения",
      y = "Значение, пункты (штриховая линия — 100)")
time_plot <- function(columns, labels, units) {
 z <- data.frame()
 for (j in seq_along(columns)) z <- rbind(z,
   data.frame(Месяц = 1:9, Показатель = labels[j], Значение = history[[columns[j]]]))
 old <- z[z$Месяц <= 8 & !is.na(z$Значение), ]; new <- z[z$Месяц == 9, ]
 p <- ggplot(z, aes(Месяц, Значение, color = Показатель)) +
  geom_line(data = old, aes(linetype = Показатель), linewidth = .7) +
  geom_point(data = old, size = 2) + geom_point(data = new, shape = 17, size = 3.5) +
  scale_x_continuous(breaks = 1:9, labels = months, limits = c(.7, 9.3)) +
  labs(x = "Месяц 2026 года", y = units, color = "Показатель", linetype = "Показатель")
 if (nrow(old) == 0) p <- p + annotate("text", x = 4.4, y = mean(new$Значение),
  label = "Январь–август: официальные значения\nне получены; линии не построены", color = "#333333", size = 4)
 p
}
fig7 <- time_plot(c("obs", "exp"), c("Наблюдаемая", "Ожидаемая"), "Медиана, % годовых")
fig8 <- time_plot(c("ITS", "IPO", "IPN"), c("ИТС", "ИПО", "ИПН"), "Индекс, пункты")
figures <- list(fig1, fig2, fig3, fig4, fig5, fig6, fig7, fig8)
for (j in seq_along(figures)) ggsave(file.path(out_dir, "figures", paste0("fig", j, ".png")),
 figures[[j]], width = ifelse(j == 5, 11, 9), height = ifelse(j == 3, 7, 5.2),
 units = "in", dpi = 180, bg = "white")
tables <- list(Качество = quality, Пригодность = usable, Пропуски = audit,
 Размер_семьи = family_stats, Частоты_Q12 = family_f,
 Перекодировка = data.frame(Категория = code_labels, Код = code_values, L = lower, h = width),
 Инфляция = infl_stats, Частоты_obs = gobs$table, Частоты_exp = gexp$table,
 Коды_дохода = data.frame(Категория = income_levels, Рубли = income_codes),
 Доход = income_stats, Ряд_дохода = income_series,
 Частные_индексы = private, Составные = composite, Группы_дохода = groups,
 Q5_по_доходу = group_q5, ИПН_по_полу = sex_ipn, Индексы_по_полу = sex_private,
 История = history, Динамика = delta)
for (nm in names(freq)) tables[[paste0("Частоты_", ifelse(nm %in% attrs, match(nm, attrs), nm))]] <- freq[[nm]]
for (nm in names(freq_all)) tables[[paste0(nm, "_от_всех")]] <- freq_all[[nm]]
openxlsx::write.xlsx(tables, file.path(out_dir, "tables.xlsx"), overwrite = TRUE)
write.csv(Data, file.path(out_dir, "calculated_data.csv"), row.names = FALSE, na = "", fileEncoding = "UTF-8")
capture.output(sessionInfo(), file = file.path(out_dir, "sessionInfo.txt"))
fmt <- function(x, digits = 1) format(round(x, digits), nsmall = digits, trim = TRUE, decimal.mark = ",")
table_no <- 0
tbl <- function(d, title, digits = 1) {
 table_no <<- table_no + 1
 cat("\n\n::: {custom-style=\"Table Caption\"}\nТаблица ", table_no, ". ", title, "\n:::\n\n", sep = "")
 for (j in seq_along(d)) if (is.numeric(d[[j]]) && !is.integer(d[[j]])) d[[j]] <- round(d[[j]], digits)
 print(knitr::kable(d, format = "pipe", row.names = FALSE, digits = digits))
 cat("\n\n")
}
median_sub <- function(gm, what) {
 if (gm$defined) {
  cat(sprintf("**%s.** Медианная категория — %s; N = %s, N/2 = %s, L = %s, h = %s, F_prev = %s, f = %s. ",
    what, gm$label, gm$N, fmt(gm$N/2), fmt(gm$L), fmt(gm$h), gm$prev, gm$f))
  cat(sprintf("Подстановка: \\(\\tilde{x}_{0.5} = %s + %s \\cdot (%s - %s)/%s = %s\\) %% годовых.\n\n",
   fmt(gm$L), fmt(gm$h), fmt(gm$N/2), gm$prev, gm$f, fmt(gm$value, 4)))
 } else cat(what, ": медианная категория открытая или точечная; интерполяция не определена. Код = ", gm$value, ".\n\n", sep = "")
}


tbl(quality, "Проверка структуры, значений и логики", 0)
tbl(usable, "Число пригодных наблюдений по показателям", 0)
tbl(audit[, 1:5], "Основания частотных таблиц, технические и недопустимые ответы", 0)
tbl(audit[, c(1, 6:8)], "Структурные пропуски и согласованность условий задавания", 0)

for (nm in c(paste0("Q", c(1:11, 13)), attrs[attrs %in% names(freq)])) {
 base <- audit$Основание[match(nm, audit$Поле)]
 tbl(freq[[nm]], paste0(nm, ": частоты и доли; основание N = ", base), 0)
}
for (nm in names(freq_all)) tbl(freq_all[[nm]], paste0(nm, ": доли от всей выборки, N = ", N), 0)
tbl(family_stats, "Q12: описательные показатели размера семьи")
tbl(family_f, "Q12: распределение числа членов семьи", 0)

tbl(data.frame(Категория = code_labels, `Код, %` = code_values,
 `Нижняя граница L` = lower, `Ширина h` = width, check.names = FALSE),
 "Правила числовой перекодировки и интерполяционные границы")
tbl(infl_stats, "Число ответов, обычные медианы, квартили и уточненные медианы, % годовых")

tbl(gobs$table, "Наблюдаемая инфляция: частоты n и накопленные частоты F", 1)
tbl(gexp$table, "Ожидаемая инфляция: частоты n и накопленные частоты F", 1)
median_sub(gobs, "Наблюдаемая инфляция")
median_sub(gexp, "Ожидаемая инфляция")

tbl(data.frame(Категория = income_levels, `Доход семьи, руб.` = income_codes,
 check.names = FALSE), "Перекодировка Q13 в совокупный доход семьи", 0)
tbl(income_stats, "Описательные показатели дохода на человека, руб. в месяц")
tbl(income_series, "Интервальный ряд дохода с шагом 5000 руб., частоты n и накопленные F", 0)

tbl(private, "Частные индексы: числители, основания, доли и итоговые значения")

tbl(composite, "Составные индексы и независимая проверка ИПН, пункты")

tbl(groups, "Границы квантильных групп дохода, основания и обычные медианы infl_exp")
tbl(group_q5, "Частоты и доли ответов Q5 по пяти группам дохода", 0)

tbl(sex_ipn, "ИПН мужчин и женщин, пункты; основания каждого частного индекса")

if (isTRUE(getOption("knitr.in.progress"))) print(fig1)

if (isTRUE(getOption("knitr.in.progress"))) print(fig2)

if (isTRUE(getOption("knitr.in.progress"))) print(fig3)

if (isTRUE(getOption("knitr.in.progress"))) print(fig4)

if (isTRUE(getOption("knitr.in.progress"))) print(fig5)

if (isTRUE(getOption("knitr.in.progress"))) print(fig6)

if (isTRUE(getOption("knitr.in.progress"))) print(fig7)

if (isTRUE(getOption("knitr.in.progress"))) print(fig8)

tbl(history[, 1:7], "Помесячные уровни и источник: NA означает отсутствие подтвержденного числа")
tbl(delta[, 1:6], "Динамика август → учебный сентябрь: разность уровней")
