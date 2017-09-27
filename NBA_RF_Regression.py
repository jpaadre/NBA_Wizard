import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
from sklearn.preprocessing import Imputer,LabelEncoder, OneHotEncoder,StandardScaler,PolynomialFeatures
from sklearn.cross_validation import train_test_split
from sklearn.linear_model import LinearRegression
import statsmodels.formula.api as sm
from sklearn.svm import SVR
from sklearn.tree import DecisionTreeRegressor
from sklearn.ensemble import RandomForestRegressor
from pandas import ExcelWriter
from sklearn.metrics import mean_squared_error,mean_absolute_error,explained_variance_score,r2_score

#-------------Get Dataset-------------
def getDataSet(dataset):
    df = pd.read_excel(dataset)
    return df

source = "C:\\Users\\Peter Haley\\Desktop\\Projects\\Data_Science\\Python\\NBA\\Excel\\"
dataset = getDataSet(source+'DataForModel.xlsx')
# print(dataset.head())
df = dataset.dropna()
# df = df.head(n=600)
# df2 = df[['GAMECODE','TEAM_ABBREVIATION_x','HomeIndex_x','OFF_RATING_x']]
#
# writer1 = ExcelWriter("RF_Analysis.xlsx")
# df2.to_excel(writer1,'Master')
# writer1.save()


def HomeOrAway(df):
    if df['HomeIndex_x'] == 1:
        return ('Home')
    else:
        return ('Away')

# df['HomeOrAway'] = df.apply(HomeOrAway,axis=1)

x = df[['HomeIndex_x','DaysRest_x','AvgPace_x','AvgORTG_x','AvgDRTG_x','AvgORTG_L5_x','AvgDRTG_L5_x','std_AvgORTG', 'std_AvgDRTG','std_AvgORTG_L5','std_AvgDRTG_L5'
,'DaysRest_y','AvgPace_y','AvgORTG_y','AvgDRTG_y','AvgORTG_L5_y','AvgDRTG_L5_y']].values
#
# x = df[['HomeOrAway','DaysRest_x','AvgORTG_x','std_AvgORTG_L5'
# ,'DaysRest_y','AvgPace_y','AvgDRTG_y']].values
#
# x = df[['HomeOrAway','DaysRest_x','AvgORTG_x'
# ,'DaysRest_y','AvgORTG_y','AvgDRTG_y','AvgDRTG_L5_y']].values


# x = df[['std_AvgORTG','std_AvgORTG_L5'
# ,'AvgDRTG_y','AvgDRTG_L5_y']].values

# -------------Building MLR dfs to test--------------
# x = df[['AvgORTG_x','AvgDRTG_y']].values
# x2 = df[['HomeIndex_x','AvgORTG_x','AvgDRTG_L5_y']].values
# x3 = df[['HomeIndex_x','std_AvgORTG','AvgDRTG_y']].values
# x4 = df[['HomeIndex_x','std_AvgORTG','AvgDRTG_L5_y']].values

# # x = dataset.iloc[:,1:2].values
y = df.iloc[:,20].values

#--------------- Encoding categorical data-------------

labelencoder_x = LabelEncoder()
x[:, 0] = labelencoder_x.fit_transform(x[:, 0])
onehotencoder = OneHotEncoder(categorical_features = [0])
x = onehotencoder.fit_transform(x).toarray()

# Avoiding the Dummy Variable Trap
x = x[:,1:]

# print(x,y)

# -------------Split Train and Test Data-------------

x_train, x_test, y_train, y_test = train_test_split(x,y,test_size =0.25, random_state =0)
print('Data Split')

# sc_x = StandardScaler()
# x_train = sc_x.fit_transform(x_train)
# x_test = sc_x.transform(x_test)
# sc_y = StandardScaler()
# y_train = sc_y.fit_transform(y_train)
# y_test = sc_y.fit_transform(y_test)


#
# # -------------Feature Scaling-------------
# # sc_x = StandardScaler()
# # sc_y = StandardScaler()
# # x = sc_x.fit_transform(x)
# # y = sc_y.fit_transform(y)
# # # print (x,y)
#
# #-------------Fit RF to Data Set-------------
regressor = RandomForestRegressor(n_estimators=10000, random_state=0)
regressor.fit(x_train,y_train)

# regressor = LinearRegression()
# regressor.fit(x_train, y_train)

r2 = regressor.score(x_train, y_train)




print('Regressing')
#
# #-------------Predict a new result with Random Forest-------------
y_pred = regressor.predict(x_test)
# print(y_pred)
print('$$$$$')

r2_2 = regressor.score(x_test, y_test)
mae = mean_absolute_error(y_test, y_pred)
mse = mean_squared_error(y_test, y_pred)
evs = explained_variance_score(y_test, y_pred)
# r2 = r2_score(y_test, y_pred)
print(r2)
print(r2_2)
print('MAE ', mae)
print('MSE ', mse)
print('Explained Variance ', evs)
print('R2 ', r2)

imp = regressor.feature_importances_
print(imp)

# df3 = pd.DataFrame({'Actual':y_test,'Predicted':y_pred})
# writer2 = ExcelWriter("RF_Results.xlsx")
# df3.to_excel(writer2,'Master')
# writer2.save()
# -----------------Building up MLR model-------------------

# import statsmodels.api as sm
# x = sm.add_constant(x)
# x_opt = x[:,[0,1,2,3]]
# regressor_ols = sm.OLS(endog = y, exog = x_opt).fit()
# print(regressor_ols.summary())
#
# x2 = sm.add_constant(x2)
# x_opt2 = x2[:,[0,1,2,3]]
# regressor_ols2 = sm.OLS(endog = y, exog = x_opt2).fit()
# print(regressor_ols2.summary())
#
# x3 = sm.add_constant(x3)
# x_opt3 = x3[:,[0,1,2,3]]
# regressor_ols3 = sm.OLS(endog = y, exog = x_opt3).fit()
# print(regressor_ols3.summary())
#
# x4 = sm.add_constant(x4)
# x_opt4 = x4[:,[0,1,2,3]]
# regressor_ols4 = sm.OLS(endog = y, exog = x_opt4).fit()
# print(regressor_ols4.summary())





# #-------------Visualize Results-------------
# plt.scatter(y_test,y_pred,color = 'red')
# # plt.plot(x,regressor.predict(x),color = 'orange')
# plt.title('ORTG Prediction')
# plt.xlabel('Actual')
# plt.ylabel('Predicted')
# plt.show()
