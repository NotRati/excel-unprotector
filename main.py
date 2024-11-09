import zipfile
import os
import shutil
import sys

class Unprotect:
    
    def __init__(self, args, unprotectXmlList=None):
        self.args = args
        self.folderPath = './extracted'
        if unprotectXmlList is None:
            self.unprotectXmlList = ['xl/workbook.xml', 'xl/styles.xml', 'xl/worksheets']
        else:
            self.unprotectXmlList = unprotectXmlList
    def run(self):
        self.setFilePath('C:\\Users\\User\\Desktop\\Roadmap-Someka-Excel-Template-V1-Free-Version.xlsx')
        self.moveExcelFile()
        self.zipConversion()
        self.extractZip()
        self.removeProtection()
        self.excelConversion()
        self.cleanUp()
    def getFilePath(self):
        if len(self.args) > 1:
            self.excelFilePath = self.args[1]
        else:
            print("No file path provided!")
            sys.exit(1)

    def setFilePath(self, filePath):
        self.excelFilePath = filePath

    def moveExcelFile(self):
        # Make a copy of the Excel file to work with
        self.excelFilePath = shutil.copy(self.excelFilePath, './unprotected.xlsx')

    def zipConversion(self):
        base = os.path.splitext(self.excelFilePath)[0]
        self.zipFilePath = f"{base}-copyyy.zip"
        try:
            os.remove(self.zipFilePath)
        except:
            pass
        os.rename(self.excelFilePath, self.zipFilePath)

    def extractZip(self):
        # Extract the contents of the zip (Excel file)
        with zipfile.ZipFile(self.zipFilePath) as f:
            f.extractall(self.folderPath)

    def deleteCopiedExcel(self):
        # Delete the original Excel file after copying
        os.remove(self.excelFilePath)

    def deleteZipFolder(self):
        # Delete the temporary zip file
        os.remove(self.zipFilePath)

    def removeProtection(self):
        for file in self.unprotectXmlList:
            # Check if it's a file or folder
            filePath = os.path.join(self.folderPath, file)
            
            if os.path.isfile(filePath):
                with open(filePath, 'r', encoding="utf-8") as f:
                    xmlContent = f.read()

                # Replace protection-related tags with new values or remove them
                xmlContent = xmlContent.replace('workbookProtection', 'wukkabookProtection')
                xmlContent = xmlContent.replace('applyProtection="1"', 'applyProtection="0"')
                xmlContent = xmlContent.replace('hidden="1"', 'hidden="0"')
                xmlContent = xmlContent.replace('locked="1"', 'locked="0"')

                with open(filePath, 'w', encoding='utf-8') as f:
                    f.write(xmlContent)
            elif os.path.isdir(filePath):
                for folderFile in os.listdir(filePath):
                    folderFilePath = os.path.join(filePath, folderFile)
                    if folderFilePath.endswith('.xml'):
                        with open(folderFilePath, 'r', encoding='utf-8') as f:
                            xmlContent = f.read()

                        # Modify XML content to remove protection tags

                        xmlContent = xmlContent.replace('workbookProtection', 'wukkabookProtection')
                        xmlContent = xmlContent.replace('sheetProtection', 'shitProtection')

                        with open(folderFilePath, 'w', encoding='utf-8') as f:
                            f.write(xmlContent)

    def excelConversion(self):
        # Create a new zip file from the modified folder
        self.zipFilePath = shutil.make_archive(self.zipFilePath.replace('.zip', ''), 'zip', self.folderPath)
        os.rename(self.zipFilePath, self.zipFilePath.replace('.zip', '.xlsx'))

    def cleanUp(self):
        # Remove the extracted folder after the conversion
        shutil.rmtree(self.folderPath)

# Initialize and run the unprotection process
unprotect = Unprotect(sys.argv)
unprotect.run()
