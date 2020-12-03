# load library to import excel data
library(readxl)

### import excel data ###
# import excel data "ruwe data"
raw <- read_excel(./data/raw/"RCT-DD-wound.xlsx", sheet = 2)

# import excel data "mscore-treated-excluded"
mscore <- read_excel("RCT-DD-wound.xlsx", sheet = 4)

# import excel data "mscore-treated-excluded-followup"
fup <- read_excel("RCT-DD-wound.xlsx", sheet = 5)

# import excel data "wound-healing-progress"
whp <- read_excel("RCT-DD-wound.xlsx", sheet = 8)



### study flow ###
#HV
names(raw) # variabele namen
summary(raw) # samenvatting per variabele

# activeer dataset "raw"
attach(raw)
# ruwe data - aantal poten gescoord, M-scores en behandelingen
(tab <- table(M_D0, Treatment,useNA = "ifany")) 
#HV eerste ( en laatste ) in voorgaande regel zorgen er voor dat het resultaat wordt getoond
#HV is in dit geval niet nodig want vervolgens doe je addmargins(..) en wordt ook dat resultaat getoond
addmargins(tab)
round(prop.table(tab, 1), 3)
rm(tab)

# ruwe data - prevalentie Mscores per bedrijf
(tab <- table(Farm, M_D0, useNA = "ifany"))
addmargins(tab)
round(prop.table(tab, 1), 3)
rm(tab)

# figuur maken met in Y-as prevalentie op D0 als percentage, X-as Farm en voor iedere farm 100% stacked column met M0, active lesions (M1 + M2 + M4.1) en chronic lesions (M3 + M4)

# de-activeer dataset "raw"
detach(raw)

# selecteer treated cases

# excludeer lost cases (dit is op basis van remarks student, vb wilde koe dus niet te vangen of verband kwam er eerder af, en bekijken ruwe dataset - hoe doe ik dit netjes in R? of is het toch gebruikelijk dit in en copy van de ruwe dataset in excel te doen?)
# geef deze dataset de nieuwe naam "mscore"

# activeer dataset "mscore"
attach(mscore)

# mscore - aantal poten gescoord, M-scores en behandelingen
(tab <- table(M_D0, Treatment,useNA = "ifany"))
addmargins(tab)
round(prop.table(tab, 1), 3)
rm(tab)

# mscore - aantal poten in iedere behandelgroep
(tab <- table(M_D0, Treatment,useNA = "ifany"))
addmargins(tab)
round(prop.table(tab, 1), 3)
rm(tab)

# mscore - aantal koeien in iedere behandelgroep

# mscore - aantal koeien met 2 poten en aantal koeien met 1 poot in de studie

# de-activeer dataset "mscore"
detach(mscore)

# activeer dataset "fup"
attach(fup)

# follow-up - aantal poten start (dag 0) in iedere behandelgroep
(tab <- table(M_D0, Treatment,useNA = "ifany"))
addmargins(tab)
round(prop.table(tab, 1), 3)
rm(tab)

# follow-up - aantal poten dag 21 in iedere behandelgroep
(tab <- table(M_D21, Treatment,useNA = "ifany"))
addmargins(tab)
round(prop.table(tab, 1), 3)
rm(tab)

# follow-up - aantal poten dag 35 in iedere behandelgroep
(tab <- table(M_D35, Treatment,useNA = "ifany"))
addmargins(tab)
round(prop.table(tab, 1), 3)
rm(tab)

# de-activeer dataset "fup"
detach(fup)



### Treatment outcomes ###

# activeer dataset "mscore"
attach(mscore)

# mscore - transitiematrix D0-D10
(tab <- table(M_D10, Treatment, M_D0, useNA = "ifany"))
addmargins(tab)
round(prop.table(tab, 1), 3)
addmargins(tab)
rm(tab)

# mscore - transitiematrix per farm D0-D10
(tab <- table(M_D10, Treatment, Farm, M_D0, useNA = "ifany"))
addmargins(tab)
round(prop.table(tab, 1), 3)
addmargins(tab)
rm(tab)

# mscore - create variable 'clinical improvement on D10 (imp_D10)'

# de-activeer dataset "mscore"
detach(mscore)

# activeer dataset "fup"
attach(fup)

# follow-up - transitiematrix D0-D21
(tab <- table(M_D21, Treatment, M_D0))
addmargins(tab)
round(prop.table(tab, 1), 3)
addmargins(tab)
#HV toevoegen proporties per subtabel
prop.table(tab, c(1,3)) # de 1 geeft aan totaliseren in de regel en de 3 betekent per M_D0
rm(tab)

# follow-up - transitiematrix D0-D35
(tab <- table(M_D35, Treatment, M_D0))
addmargins(tab)
round(prop.table(tab, 1), 3)
addmargins(tab)
rm(tab)

# fup - create variable 'clinical improvement on D21 (imp_D21)'

# fup - pearson chi squared test of independence for treatment and clinical improvement on D21
chisq.test(Treatment, imp_D21)

# fup - create variable 'clinical improvement on D35 (imp_D35)'

# fup - pearson chi squared test of independence for treatment and clinical improvement on D35
chisq.test(Treatment, imp_D35)

# de-activeer dataset "fup"
detach(fup)

# activeer dataset "whp"
attach(whp)

# whp - overview wound healing progress outcome D0-D10 per treatment group
(tab <- table(WHP_D010, Treatment, useNA = "ifany"))
addmargins(tab)
round(prop.table(tab, 1), 3)
rm(tab)

# whp - overview wound healing progress outcome D0-D3 per treatment group
(tab <- table(WHP_D03, Treatment, useNA = "ifany"))
addmargins(tab)
round(prop.table(tab, 1), 3)
rm(tab)

# whp - overview wound healing progress outcome D3-D7 per treatment group
(tab <- table(WHP_D37, Treatment, useNA = "ifany"))
addmargins(tab)
round(prop.table(tab, 1), 3)
rm(tab)

# whp - overview wound healing progress outcome D7-D10 per treatment group
(tab <- table(WHP_D710, Treatment, useNA = "ifany"))
addmargins(tab)
round(prop.table(tab, 1), 3)
rm(tab)

# whp - overview wound healing progress outcome D0-D10 per treatment group for each D0 M-score
(tab <- table(WHP_D010_rec, Treatment, M_D0, useNA = "ifany"))
addmargins(tab)
round(prop.table(tab, 1), 3)
rm(tab)

# de-activeer dataset "whp"
detach(whp)



###### Analyse ######
#load library linear mixed effects (multivariabele logistische regressie met farm als random effect)
library(lme4)

### UNIVARIABLE logistische regressie met farm als random effect ###

## M-score ##

# univariable logistische regressie clinical improvement mscore D10 & treatment group
fit <- glmer(imp_D10 ~ Treatment + (1|Farm), family="binomial", data=mscore) 
#HV  doe eens: table(mscore$imp_D10, mscore$Farm, mscore$Treatment)
# dan zie je dat in de de aantallen erg laag zijn met score 0 -> probleem met schatten
summary(fit) 
drop1(fit, test= "Chisq")
beta <- exp(getME(fit, "beta"))
ci <- exp(confint(fit))
round(cbind(beta, ci[2:3,]), 2)
rm(fit, beta, ci)

# univariable logistische regressie clinical improvement mscore D10 & Mscore D0
fit <- glmer(imp_D10 ~ factor(M_D0) + (1|Farm), family="binomial", data=mscore) 
#HV  doe eens: table(mscore$imp_D10, mscore$Farm, mscore$M_D0)
summary(fit) 
drop1(fit, test= "Chisq")
beta <- exp(getME(fit, "beta"))
ci <- exp(confint(fit))
round(cbind(beta, ci[2:4,]), 2)
rm(fit, beta, ci)

# univariable logistische regressie clinical improvement mscore D10 & time under bandage
fit <- glmer(imp_D10 ~ Bandage_rec + (1|Farm), family="binomial", data=mscore) 
summary(fit)
drop1(fit, test= "Chisq")
beta <- exp(getME(fit, "beta"))
ci <- exp(confint(fit))
round(cbind(beta, ci[2:3,]), 2)
rm(fit, beta, ci)



## wound healing progress met unable to score als 'not improved' ##

# univariable logistische regressie wound healing progress between D0 and D10 & treatment group
fit <- glmer(WHP_D010_rec0 ~ Treatment + (1|Farm), family="binomial", data=whp) 
summary(fit) 
drop1(fit, test= "Chisq")
beta <- exp(getME(fit, "beta"))
ci <- exp(confint(fit))
round(cbind(beta, ci[2:3,]), 2)
rm(fit, beta, ci)

# univariable logistische regressie wound healing progress between D0 and D10 & Mscore D0
fit <- glmer(WHP_D010_rec0 ~ factor(M_D0) + (1|Farm), family="binomial", data=whp) 
summary(fit) 
drop1(fit, test= "Chisq")
beta <- exp(getME(fit, "beta"))
ci <- exp(confint(fit))
round(cbind(beta, ci[2:4,]), 2)
rm(fit, beta, ci)

# univariable logistische regressie wound healing progress between D0 and D10 & time under bandage
fit <- glmer(WHP_D010_rec0 ~ Bandage_rec + (1|Farm), family="binomial", data=whp) 
summary(fit) 
drop1(fit, test= "Chisq")
beta <- exp(getME(fit, "beta"))
ci <- exp(confint(fit))
round(cbind(beta, ci[2:3,]), 2)
rm(fit, beta, ci)

# univariable logistische regressie wound healing progress between D0 and D10 & wound healing progress between D0 and D3
fit <- glmer(WHP_D010_rec0 ~ WHP_D03_rec0 + (1|Farm), family="binomial", data=whp) 
summary(fit)
#HV in de drop1 wordt een warning gegeven. Error in drop1.merMod(fit, test = "Chisq") : 
# number of rows in use has changed: remove missing values?
# dit betekenis dat wrs WHP_D03_rec0 een missing heeft zie summary(whp) of summary(whp$WHP_D03_rec0)
# inderdaad 1 NA, dus het model inclusief en exclusief deze variabele zijn gebaseerd op een
# verschillend aantal records.
fit <- glmer(WHP_D010_rec0 ~ WHP_D03_rec0 + (1|Farm), family="binomial", data=whp[!is.na(whp$WHP_D03_rec0),])
#HV met data=whp[!is.na(whp$WHP_D03_rec0),] worden alleen de records geselecteerd zonder NA in deze variabele
summary(fit)
drop1(fit, test= "Chisq")
beta <- exp(getME(fit, "beta"))
ci <- exp(confint(fit))
round(cbind(beta, ci[2:3,]), 2)
rm(fit, beta, ci)

# univariable logistische regressie wound healing progress between D0 and D10 & wound healing progress between D3 and D7
#HV toegevoegd: data=whp[!is.na(whp$WHP_D37_rec0),])
fit <- glmer(WHP_D010_rec0 ~ WHP_D37_rec0 + (1|Farm), family="binomial", data=whp[!is.na(whp$WHP_D37_rec0),]) 
summary(fit) 
drop1(fit, test= "Chisq")
beta <- exp(getME(fit, "beta"))
ci <- exp(confint(fit))
round(cbind(beta, ci[2:3,]), 2)
rm(fit, beta, ci)

# univariable logistische regressie wound healing progress between D0 and D10 & wound healing progress between D7 and D10
fit <- glmer(WHP_D010_rec0 ~ WHP_D710_rec0 + (1|Farm), family="binomial", data=whp) 
summary(fit) 
drop1(fit, test= "Chisq")
beta <- exp(getME(fit, "beta"))
ci <- exp(confint(fit))
round(cbind(beta, ci[2:3,]), 2)
rm(fit, beta, ci)



## wound healing progress met unable to score als 'improved' ##

# univariable logistische regressie wound healing progress between D0 and D10 & treatment group
fit <- glmer(WHP_D010_rec1 ~ Treatment + (1|Farm), family="binomial", data=whp) 
summary(fit) 
drop1(fit, test= "Chisq")
beta <- exp(getME(fit, "beta"))
ci <- exp(confint(fit))
round(cbind(beta, ci[2:3,]), 2)
rm(fit, beta, ci)

# univariable logistische regressie wound healing progress between D0 and D10 & Mscore D0
fit <- glmer(WHP_D010_rec1 ~ factor(M_D0) + (1|Farm), family="binomial", data=whp) 
summary(fit) 
drop1(fit, test= "Chisq")
beta <- exp(getME(fit, "beta"))
ci <- exp(confint(fit))
round(cbind(beta, ci[2:4,]), 2)
rm(fit, beta, ci)

# univariable logistische regressie wound healing progress between D0 and D10 & time under bandage
fit <- glmer(WHP_D010_rec1 ~ Bandage_rec + (1|Farm), family="binomial", data=whp) 
summary(fit) 
drop1(fit, test= "Chisq")
beta <- exp(getME(fit, "beta"))
ci <- exp(confint(fit))
round(cbind(beta, ci[2:3,]), 2)
rm(fit, beta, ci)

# univariable logistische regressie wound healing progress between D0 and D10 & wound healing progress between D0 and D3
fit <- glmer(WHP_D010_rec1 ~ WHP_D03_rec1 + (1|Farm), family="binomial", data=whp) 
summary(fit) 
drop1(fit, test= "Chisq")
beta <- exp(getME(fit, "beta"))
ci <- exp(confint(fit))
round(cbind(beta, ci[2:3,]), 2)
rm(fit, beta, ci)

# univariable logistische regressie wound healing progress between D0 and D10 & wound healing progress between D3 and D7
fit <- glmer(WHP_D010_rec1 ~ WHP_D37_rec1 + (1|Farm), family="binomial", data=whp) 
summary(fit) 
drop1(fit, test= "Chisq")
beta <- exp(getME(fit, "beta"))
ci <- exp(confint(fit))
round(cbind(beta, ci[2:3,]), 2)
rm(fit, beta, ci)

# univariable logistische regressie wound healing progress between D0 and D10 & wound healing progress between D7 and D10
fit <- glmer(WHP_D010_rec1 ~ WHP_D710_rec1 + (1|Farm), family="binomial", data=whp) 
summary(fit) 
drop1(fit, test= "Chisq")
beta <- exp(getME(fit, "beta"))
ci <- exp(confint(fit))
round(cbind(beta, ci[2:3,]), 2)
rm(fit, beta, ci)



### MULTIVARIABLE linear mixed effects model, with farm as random effect, manual backwards elimination based on lowest AIC ###

## M-score ##
#full LME model
#HV omdat in de univariabele modellen er al een probleem was met de lage aantallen, zal zeker met 
# multivariabele modellen het probleem alleen maar groter worden!
fit <- glmer(imp_D10 ~ Treatment + factor(M_D0) + Bandage_rec + (1|Farm), family="binomial", data=mscore) 
summary(fit)
drop1(fit, test= "Chisq")
beta <- exp(getME(fit, "beta"))
ci <- exp(confint(fit))
round(cbind(beta, ci[2:6,]), 2)
rm(fit, beta, ci)
# Alle single term deletion AIC zijn hoger dan AIC model en we willen eigenlijk een lagere AIC
# Minst hoge AIC tov model is voor M-score D0: delta AIC = 1.99 dit is kleiner dan |2| dus M-score D0 uit model halen omdat je zo minder variabelen nodig hebt voor je model, ook al past het iets minder goed (omdat de AIC groter wordt)
# LME  model zonder (M-score D0)
fit <- glmer(imp_D10 ~ Treatment + Bandage_rec + (1|Farm), family="binomial", data=mscore) 
summary(fit)
drop1(fit, test= "Chisq")
beta <- exp(getME(fit, "beta"))
ci <- exp(confint(fit))
round(cbind(beta, ci[2:4,]), 2)
rm(fit, beta, ci)
# AIC model is van 151.0 naar 153.0 gestegen, binnen grens van |2| dus ok ondanks slechtere fit
# Alle single term deletion AIC zijn groter dan model en we willen eigenlijk een lagere AIC
# Minst hoge AIC tov model is voor Bandage: delta AIC = 24.02 dit is groter dan |2| dus geen variabelen meer uit model halen



## wound healing progress met unable to score als 'not improved' ##
#full LME model
fit <- glmer(WHP_D010_rec0 ~ Treatment + factor(M_D0) + Bandage_rec + WHP_D03_rec0 + WHP_D37_rec0 + WHP_D710_rec0 + (1|Farm), family="binomial", data=whp) 
summary(fit)
drop1(fit, test= "Chisq")
beta <- exp(getME(fit, "beta"))
ci <- exp(confint(fit))
round(cbind(beta, ci[2:9,]), 2)
rm(fit, beta, ci)
# variable met lowest single term deletion AIC is M-score D0 en lager dan AIC model
# LME  model zonder (M-score D0)
fit <- glmer(WHP_D010_rec0 ~ Treatment + Bandage_rec + WHP_D03_rec0 + WHP_D37_rec0 + WHP_D710_rec0 + (1|Farm), family="binomial", data=whp) 
summary(fit)
drop1(fit, test= "Chisq")
beta <- exp(getME(fit, "beta"))
ci <- exp(confint(fit))
round(cbind(beta, ci[2:7,]), 2)
rm(fit, beta, ci)
# variable met lowest single term deletion AIC is Bandage en lager dan AIC model
# LME  model zonder (M-score D0 en Bandage)
fit <- glmer(WHP_D010_rec0 ~ Treatment + WHP_D03_rec0 + WHP_D37_rec0 + WHP_D710_rec0 + (1|Farm), family="binomial", data=whp) 
summary(fit)
drop1(fit, test= "Chisq")
beta <- exp(getME(fit, "beta"))
ci <- exp(confint(fit))
round(cbind(beta, ci[2:6,]), 2)
rm(fit, beta, ci)
# Alle single term deletion AIC zijn hoger dan AIC model en we willen eigenlijk een lagere AIC
# Minst hoge AIC tov model is voor WP0-3: delta AIC = 0.51 dit is kleiner dan |2| dus WHP0-3 uit model halen omdat je zo minder variabelen nodig hebt voor je model, ook al past het iets minder goed (omdat de AIC groter wordt)
# LME  model zonder (M-score D0, Bandage en WHP0-3)
fit <- glmer(WHP_D010_rec0 ~ Treatment + WHP_D37_rec0 + WHP_D710_rec0 + (1|Farm), family="binomial", data=whp) 
summary(fit)
drop1(fit, test= "Chisq")
beta <- exp(getME(fit, "beta"))
ci <- exp(confint(fit))
round(cbind(beta, ci[2:5,]), 2)
rm(fit, beta, ci)
# ERROR! ik krijg geen single term deletion AIC meer en kan dus niet meer verder... :(
# komt dit omdat er voor de WHP0-3 missing values zijn? Er zijn daar nl enkele lege cellen omdat deze foto's ontbraken in de ruwe data.


## wound healing progress met unable to score als ' improved' ##
#full LME model
fit <- glmer(WHP_D010_rec1 ~ Treatment + factor(M_D0) + Bandage_rec + WHP_D03_rec1 + WHP_D37_rec1 + WHP_D710_rec1 + (1|Farm), family="binomial", data=whp) 
summary(fit)
drop1(fit, test= "Chisq")
beta <- exp(getME(fit, "beta"))
ci <- exp(confint(fit))
round(cbind(beta, ci[2:9,]), 2)
rm(fit, beta, ci)
# variable met lowest single term deletion AIC is M-score D0 en lager dan AIC model
# LME  model zonder (M-score D0)
fit <- glmer(WHP_D010_rec1 ~ Treatment + Bandage_rec + WHP_D03_rec1 + WHP_D37_rec1 + WHP_D710_rec1 + (1|Farm), family="binomial", data=whp) 
summary(fit)
drop1(fit, test= "Chisq")
beta <- exp(getME(fit, "beta"))
ci <- exp(confint(fit))
round(cbind(beta, ci[2:7,]), 2)
rm(fit, beta, ci)
# variable met lowest single term deletion AIC is WHP3-7 en lager dan AIC model
# LME  model zonder (M-score D0 en WHP3-7)
fit <- glmer(WHP_D010_rec1 ~ Treatment + Bandage_rec + WHP_D03_rec1 + WHP_D710_rec1 + (1|Farm), family="binomial", data=whp) 
summary(fit)
drop1(fit, test= "Chisq")
beta <- exp(getME(fit, "beta"))
ci <- exp(confint(fit))
round(cbind(beta, ci[2:6,]), 2)
rm(fit, beta, ci)
# ERROR! ik krijg geen single term deletion AIC meer en kan dus niet meer verder... :(
# komt dit omdat er voor de WHP3-7 missing values zijn? Er zijn daar nl enkele lege cellen omdat deze foto's ontbraken in de ruwe data.