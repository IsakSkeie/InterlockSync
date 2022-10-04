#%%
import pandas as pd



class Data:
    
    def __init__(self):    
        self.FileNameExcel = "Interlock.xlsx"
        self.FileNameCSV = "GeneralTemplate.csv"
        self.SheetName = "Function In Use"
        


    def LoadInterlockExcel(self):
        #add try statement 
        self.ExcelData = pd.read_excel(self.FileNameExcel, self.SheetName)

    def LoadCSV(self):
        self.GeneralTemplate = pd.read_csv(self.FileNameCSV)

    def LoadTemplates(self):
        d = {"Templates":[":TEMPLATE=$SBE_FR", ":TEMPLATE=$SBET_FR", ":TEMPLATE=$SB_FR", ":TEMPLATE=$SBV_FR"]}
        self.Templates = pd.DataFrame(data = d)


    def ExcelToDF(self, excel, template):
        export = template
        for i, row in excel.iterrows():
            template.loc[i] = [row["Tag"], row["XDWN"], row["XINT1 Comment"], row["XINT2 Comment"], row["XINT3 Comment"],0 ,0,0,0,0]
        self.InterlockExport = template
    
    def WriteDataframeCSV(self):
        for i, temp in self.Templates.iterrows():
            with open("output.csv", 'a') as f:
                f.write("\n" + temp["Templates"] + "\n")
            
            self.InterlockExport.to_csv("Output.csv", index=False, mode='a')
            
#%%
if __name__ == "__main__":

    
    data = Data()
    data.LoadInterlockExcel()
    data.LoadCSV()
    data.LoadTemplates()

    excel = data.ExcelData
    template = data.GeneralTemplate

    data.ExcelToDF(excel, template)
    
    data.WriteDataframeCSV()
    templates = data.Templates
# %%
