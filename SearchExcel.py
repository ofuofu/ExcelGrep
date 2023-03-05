import openpyxl
import re
from pathlib import Path
from logging import getLogger,StreamHandler,INFO,DEBUG,WARN,ERROR,CRITICAL
from logging import Formatter
from ExcelGrepLogger import ExcelGrepLogger
from SearchResult import SearchResult
from AppWarnException import AppWarnException

class SearchExcel:    
    def __init__(self):
        self.logger = ExcelGrepLogger()

    def search(self, path, keyword1, keyword2, condition):
        # Pathオブジェクト生成
        targetPath = Path(path)
        
        # 検索対象の拡張子を正規表現で指定
        pattern = re.compile(r'\.xlsx|\.xlsm$')
        
        # 存在チェック
        if not targetPath.exists():            
            # print('Pathが存在しません。')
            # return
            self.logger.dubug('Pathが存在しません。')
            raise AppWarnException('Pathが存在しません。')
            
        # '**'⇒フォルダとサブフォルダ
        self.resultList = list()
        for file in targetPath.glob('**/*'):
            # self.logger.debug(str(file))
            if file.is_file() and pattern.search(str(file)):
                self.excelSearch(str(file), keyword1, keyword2, condition)
        
        return self.resultList

    def excelSearch(self, bookPath, keyword1, keyword2, condition):
        wbook = openpyxl.load_workbook(bookPath, read_only = True, data_only = True, keep_links = False)
        
        # シート単位に情報を表示するため、見つかったら次のシートを検索
        isFoundKeyword = False
        for wsheet in wbook.worksheets :
            isFoundKeyword = False
            for cells in wsheet.rows:
                # 同一のシートの中ですでに見つかった場合は、Breakする。
                if isFoundKeyword :
                    break
                
                for cell in cells:
                    if cell.value is not None:
                        # 入力のあるセル
                        try:
                            # セルのデータを文字列に変換
                            value = str(cell.value)
                        # 文字列に変換できないデータはスキップ
                        except:
                            continue
                        
                        if condition == 0 and keyword2 == "" :
                            # キーワードがセルの文字に含まれるか判定
                            if keyword1 in value:
                                result = SearchResult(bookPath, wsheet.title)
                                self.resultList.append(result)
                                isFoundKeyword = True
                                if isFoundKeyword :
                                    break

                        elif  condition == 0 and keyword2 != "" :
                            # キーワードがセルの文字に含まれるか判定
                            if keyword1 in value:
                                if keyword2 in value:
                                    result = SearchResult(bookPath, wsheet.title)
                                    self.resultList.append(result)
                                    isFoundKeyword = True
                                    if isFoundKeyword :
                                        break                                    
        wbook.close()            
        