#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
import glob
import os, sys

def main():
	src_path = os.path.join(os.path.dirname(sys.path[0]), "src")
	os.chdir(src_path)
	ods = glob.glob("*.ods")[0]
	print(ods)




if __name__ == "__main__":
	main()



# def macro(documentevent=None):  # 引数は文書のイベント駆動用。  
# 	pass
# 
# 
# 
# 
# g_exportedScripts = macro, #マクロセレクターに限定表示させる関数をタプルで指定。		
# if __name__ == "__main__":  # オートメーションで実行するとき
# 	def automation():  # オートメーションのためにglobalに出すのはこの関数のみにする。
# 		import officehelper
# 		from functools import wraps
# 		import sys
# 		from com.sun.star.beans import PropertyValue  # Struct
# 		from com.sun.star.script.provider import XScriptContext  
# 		def connectOffice(func):  # funcの前後でOffice接続の処理
# 			@wraps(func)
# 			def wrapper():  # LibreOfficeをバックグラウンドで起動してコンポーネントテクストとサービスマネジャーを取得する。
# 				try:
# 					ctx = officehelper.bootstrap()  # コンポーネントコンテクストの取得。
# 				except:
# 					print("Could not establish a connection with a running office.", file=sys.stderr)
# 					sys.exit()
# 				print("Connected to a running office ...")
# 				smgr = ctx.getServiceManager()  # サービスマネジャーの取得。
# 				print("Using {} {}".format(*_getLOVersion(ctx, smgr)))  # LibreOfficeのバージョンを出力。
# 				return func(ctx, smgr)  # 引数の関数の実行。
# 			def _getLOVersion(ctx, smgr):  # LibreOfficeの名前とバージョンを返す。
# 				cp = smgr.createInstanceWithContext('com.sun.star.configuration.ConfigurationProvider', ctx)
# 				node = PropertyValue(Name = 'nodepath', Value = 'org.openoffice.Setup/Product' )  # share/registry/main.xcd内のノードパス。
# 				ca = cp.createInstanceWithArguments('com.sun.star.configuration.ConfigurationAccess', (node,))
# 				return ca.getPropertyValues(('ooName', 'ooSetupVersion'))  # LibreOfficeの名前とバージョンをタプルで返す。
# 			return wrapper
# 		@connectOffice  # createXSCRIPTCONTEXTの引数にctxとsmgrを渡すデコレータ。
# 		def createXSCRIPTCONTEXT(ctx, smgr):  # XSCRIPTCONTEXTを生成。
# 			class ScriptContext(unohelper.Base, XScriptContext):
# 				def __init__(self, ctx):
# 					self.ctx = ctx
# 				def getComponentContext(self):
# 					return self.ctx
# 				def getDesktop(self):
# 					return ctx.getByName('/singletons/com.sun.star.frame.theDesktop')  # com.sun.star.frame.Desktopはdeprecatedになっている。
# 				def getDocument(self):
# 					return self.getDesktop().getCurrentComponent()
# 			return ScriptContext(ctx)  
# 		XSCRIPTCONTEXT = createXSCRIPTCONTEXT()  # XSCRIPTCONTEXTの取得。
# 		doc = XSCRIPTCONTEXT.getDocument()  # 現在開いているドキュメントを取得。
# 		doctype = "scalc", "com.sun.star.sheet.SpreadsheetDocument"  # Calcドキュメントを開くとき。
# 	# 	doctype = "swriter", "com.sun.star.text.TextDocument"  # Writerドキュメントを開くとき。
# 		if (doc is None) or (not doc.supportsService(doctype[1])):  # ドキュメントが取得できなかった時またはCalcドキュメントではない時
# 			XSCRIPTCONTEXT.getDesktop().loadComponentFromURL("private:factory/{}".format(doctype[0]), "_blank", 0, ())  # ドキュメントを開く。ここでdocに代入してもドキュメントが開く前にmacro()が呼ばれてしまう。
# 		flg = True
# 		while flg:
# 			doc = XSCRIPTCONTEXT.getDocument()  # 現在開いているドキュメントを取得。
# 			if doc is not None:
# 				flg = (not doc.supportsService(doctype[1]))  # ドキュメントタイプが確認できたらwhileを抜ける。
# 		return XSCRIPTCONTEXT
# 	XSCRIPTCONTEXT = automation()  # XSCRIPTCONTEXTを取得。	
# 	macro()  # マクロの実行。