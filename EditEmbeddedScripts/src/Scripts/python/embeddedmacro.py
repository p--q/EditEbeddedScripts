#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
import os
from com.sun.star.awt import XEnhancedMouseClickHandler
from com.sun.star.awt import MouseButton  # 定数
from com.sun.star.ui import XContextMenuInterceptor
from com.sun.star.ui.ContextMenuInterceptorAction import EXECUTE_MODIFIED  # enum
from com.sun.star.ui import ActionTriggerSeparatorType  # 定数
try:
	from fordebugging import enableRemoteDebugging  # デバッグ用。マクロで実行した時。
except:
	from fordebugging.pythonpath import enableRemoteDebugging  # マクロ以外の時。
# def enableRemoteDebugging(func):  # デバッグサーバーに接続したい関数やメソッドにつけるデコレーター。主にリスナーのメソッドのデバッグ目的。ただしマウスハンドラはフリーズするので直接pydevを書き込んだほうがよい。
# 	def wrapper(*args, **kwargs):
# 		frame = None
# 		doc = XSCRIPTCONTEXT.getDocument()
# 		if doc:  # ドキュメントが取得できた時
# 			frame = doc.getCurrentController().getFrame()  # ドキュメントのフレームを取得。
# 		else:
# 			currentframe = XSCRIPTCONTEXT.getDesktop().getCurrentFrame()  # モードレスダイアログのときはドキュメントが取得できないので、モードレスダイアログのフレームからCreatorのフレームを取得する。
# 			frame = currentframe.getCreator()
# 		if frame:   
# 			import time
# 			indicator = frame.createStatusIndicator()  # フレームからステータスバーを取得する。
# 			maxrange = 2  # ステータスバーに表示するプログレスバーの目盛りの最大値。2秒ロスするが他に適当な告知手段が思いつかない。
# 			indicator.start("Trying to connect to the PyDev Debug Server for about 20 seconds.", maxrange)  # ステータスバーに表示する文字列とプログレスバーの目盛りを設定。
# 			t = 1  # プレグレスバーの初期値。
# 			while t<=maxrange:  # プログレスバーの最大値以下の間。
# 				indicator.setValue(t)  # プレグレスバーの位置を設定。
# 				time.sleep(1)  # 1秒待つ。
# 				t += 1  # プログレスバーの目盛りを増やす。
# 			indicator.end()  # reset()の前にend()しておかないと元に戻らない。
# 			indicator.reset()  # ここでリセットしておかないと例外が発生した時にリセットする機会がない。
# 		import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)  # デバッグサーバーを起動していた場合はここでブレークされる。import pydevdは時間がかかる。
# 		try:
# 			func(*args, **kwargs)  # Step Intoして中に入る。	
# 		except:
# 			import traceback; traceback.print_exc()  # これがないとPyDevのコンソールにトレースバックが表示されない。stderrToServer=Trueが必須。
# 	return wrapper
def macro(documentevent=None):  # 引数は文書のイベント駆動用。  
	doc = XSCRIPTCONTEXT.getDocument() if documentevent is None else documentevent.Source  # ドキュメントのモデルを取得。 
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
	
	
	
	
	controller = doc.getCurrentController()  # コントローラの取得。
	controller.addEnhancedMouseClickHandler(EnhancedMouseClickHandler())  # マウスハンドラをコントローラに設定。
	controller.registerContextMenuInterceptor(ContextMenuInterceptor(ctx, doc))  # コントローラにContextMenuInterceptorを登録する。
class EnhancedMouseClickHandler(unohelper.Base, XEnhancedMouseClickHandler): # マウスハンドラ
	def mousePressed(self, enhancedmouseevent):  # マウスボタンをクリックした時。ブーリアンを返さないといけない。
		target = enhancedmouseevent.Target  # ターゲットを取得。
		if enhancedmouseevent.Buttons==MouseButton.LEFT:  # 左ボタンのとき
			if enhancedmouseevent.ClickCount==2:  # ダブルクリックの時
# 				import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)  
				if target.supportsService("com.sun.star.sheet.SheetCell"):  # ターゲットがセルの時。
					sheet = target.getSpreadsheet()  # ターゲットがあるシートを取得。
					celladdress = target.getCellAddress()  # ターゲットのセルアドレスを取得。
					
					
					
					return False  # セル編集モードにしない。
		return True  # Falseを返すと右クリックメニューがでてこなくなる。
	def mouseReleased(self, enhancedmouseevent):  # ブーリアンを返さないといけない。
		return True  # Trueでイベントを次のハンドラに渡す。
	def disposing(self, eventobject):
		pass	
class ContextMenuInterceptor(unohelper.Base, XContextMenuInterceptor):  # コンテクストメニューのカスタマイズ。
	def __init__(self, ctx, doc):
		self.baseurl = getBaseURL(ctx, doc)  # ScriptingURLのbaseurlを取得。
	def notifyContextMenuExecute(self, contextmenuexecuteevent):  # 右クリックで呼ばれる関数。contextmenuexecuteevent.ActionTriggerContainerを操作しないとコンテクストメニューが表示されない。
		baseurl = self.baseurl  # ScriptingURLのbaseurlを取得。
		contextmenu = contextmenuexecuteevent.ActionTriggerContainer  # コンテクストメニューコンテナの取得。
		name = contextmenu.getName().rsplit("/")[-1]  # コンテクストメニューの名前を取得。
		addMenuentry = menuentryCreator(contextmenu)  # 引数のActionTriggerContainerにインデックス0から項目を挿入する関数を取得。
		if name=="cell":  # セルのとき
			del contextmenu[:]  # contextmenu.clear()は不可。
			addMenuentry("ActionTrigger", {"CommandURL": ".uno:Cut"})
			addMenuentry("ActionTrigger", {"CommandURL": ".uno:Copy"})
			addMenuentry("ActionTrigger", {"CommandURL": ".uno:Paste"})
		elif name=="rowheader":  # 行ヘッダーのとき。
			del contextmenu[:]  # contextmenu.clear()は不可。
			addMenuentry("ActionTrigger", {"CommandURL": ".uno:Cut"})
			addMenuentry("ActionTrigger", {"CommandURL": ".uno:Copy"})
			addMenuentry("ActionTrigger", {"CommandURL": ".uno:Paste"})
			addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
			addMenuentry("ActionTrigger", {"CommandURL": ".uno:InsertRowsBefore"})
			addMenuentry("ActionTrigger", {"CommandURL": ".uno:DeleteRows"}) 
		elif name=="colheader":  # 列ヘッダーの時。
			pass  # contextmenuを操作しないとすべての項目が表示されない。
		elif name=="sheettab":  # シートタブの時。
			del contextmenu[:]  # contextmenu.clear()は不可。
			addMenuentry("ActionTrigger", {"CommandURL": ".uno:Move"})
		return EXECUTE_MODIFIED	

# class ContextMenuInterceptor(unohelper.Base, XContextMenuInterceptor):  # コンテクストメニューのカスタマイズ。
# 	def __init__(self, ctx, doc):
# 		self.baseurl = getBaseURL(ctx, doc)  # ScriptingURLのbaseurlを取得。
# 		global exportAsCSV, exportAsPDF, exportAsODS, SelectionToNewSheet   # ScriptingURLで呼び出す関数。オートメーションやAPSOでは不可。
# 		exportAsCSV, exportAsPDF, exportAsODS, SelectionToNewSheet = globalFunctionCreator(ctx, doc, sheet)  # クロージャーでScriptingURLで呼び出す関数に変数を渡す。
# 	def notifyContextMenuExecute(self, contextmenuexecuteevent):  # 引数はContextMenuExecuteEvent Struct。
# 		baseurl = self.baseurl  # ScriptingURLのbaseurlを取得。
# 		contextmenu = contextmenuexecuteevent.ActionTriggerContainer  # すでにあるコンテクストメニュー(アクショントリガーコンテナ)を取得。
# 		submenucontainer = contextmenu.createInstance("com.sun.star.ui.ActionTriggerContainer")  # サブメニューにするアクショントリガーコンテナをインスタンス化。
# 		addMenuentry(submenucontainer, "ActionTrigger", 0, {"Text": "Export as CSV...", "CommandURL": baseurl.format(exportAsCSV.__name__)})  # サブメニューを挿入。引数のない関数名を渡す。
# 		addMenuentry(submenucontainer, "ActionTrigger", 1, {"Text": "Export as PDF...", "CommandURL": baseurl.format(exportAsPDF.__name__)})  # サブメニューを挿入。引数のない関数名を渡す。
# 		addMenuentry(submenucontainer, "ActionTrigger", 2, {"Text": "Export as ODS...", "CommandURL": baseurl.format(exportAsODS.__name__)})  # サブメニューを挿入。引数のない関数名を渡す。
# 		addMenuentry(submenucontainer, "ActionTrigger", 3, {"Text": "Selection to New Sheet", "CommandURL": baseurl.format(SelectionToNewSheet.__name__)})
# 		addMenuentry(contextmenu, "ActionTrigger", 0, {"Text": "ExportAs", "SubContainer": submenucontainer})  # サブメニューを一番上に挿入。
# 		addMenuentry(contextmenu, "ActionTriggerSeparator", 1, {"SeparatorType": ActionTriggerSeparatorType.LINE})  # アクショントリガーコンテナのインデックス1にセパレーターを挿入。
		
		
		
# 		return EXECUTE_MODIFIED  # このContextMenuInterceptorでコンテクストメニューのカスタマイズを終わらす。
def menuentryCreator(menucontainer):  # 引数のActionTriggerContainerにインデックス0から項目を挿入する関数を取得。
	i = 0  # インデックスを初期化する。
	def addMenuentry(menutype, props):  # i: index, propsは辞書。menutypeはActionTriggerかActionTriggerSeparator。
		menuentry = menucontainer.createInstance("com.sun.star.ui.{}".format(menutype))  # ActionTriggerContainerからインスタンス化する。
		[menuentry.setPropertyValue(key, val) for key, val in props.items()]  #setPropertyValuesでは設定できない。エラーも出ない。
		menucontainer.insertByIndex(i, menuentry)  # submenucontainer[i]やsubmenucontainer[i:i]は不可。挿入以降のメニューコンテナの項目のインデックスは1増える。
		nonlocal i
		i += 1  # インデックスを増やす。
	return addMenuentry
def addMenuentry(menucontainer, menutype, i, props):  # i: index, propsは辞書。menutypeはActionTriggerかActionTriggerSeparator。
	menuentry = menucontainer.createInstance("com.sun.star.ui.{}".format(menutype))  # ActionTriggerContainerからインスタンス化する。
	[menuentry.setPropertyValue(key, val) for key, val in props.items()]  #setPropertyValuesでは設定できない。エラーも出ない。
	menucontainer.insertByIndex(i, menuentry)  # submenucontainer[i]やsubmenucontainer[i:i]は不可。挿入以降のメニューコンテナの項目のインデックスは1増える。
def getBaseURL(ctx, doc):	 # 埋め込みマクロ、オートメーション、マクロセレクターに対応してScriptingURLのbaseurlを返す。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
	modulepath = __file__  # ScriptingURLにするマクロがあるモジュールのパスを取得。ファイルのパスで場合分け。sys.path[0]は__main__の位置が返るので不可。
	ucp = "vnd.sun.star.tdoc:"  # 埋め込みマクロのucp。
	if modulepath.startswith(ucp):  # 埋め込みマクロの時。__file__はvnd.sun.star.tdoc:/4/Scripts/python/filename.pyというように返ってくる。
		filepath = modulepath.replace(ucp, "")  #  ucpを除去。
		transientdocumentsdocumentcontentfactory = smgr.createInstanceWithContext("com.sun.star.frame.TransientDocumentsDocumentContentFactory", ctx)
		transientdocumentsdocumentcontent = transientdocumentsdocumentcontentfactory.createDocumentContent(doc)
		contentidentifierstring = transientdocumentsdocumentcontent.getIdentifier().getContentIdentifier()  # __file__の数値部分に該当。
		macrofolder = "{}/Scripts/python".format(contentidentifierstring.replace(ucp, ""))  #埋め込みマクロフォルダへのパス。	
		location = "document"  # マクロの場所。	
	else:
		filepath = unohelper.fileUrlToSystemPath(modulepath) if modulepath.startswith("file://") else modulepath # オートメーションの時__file__はシステムパスだが、マクロセレクターから実行するとfileurlが返ってくる。
		pathsubstservice = smgr.createInstanceWithContext("com.sun.star.comp.framework.PathSubstitution", ctx)
		fileurl = pathsubstservice.substituteVariables("$(user)/Scripts/python", True)  # $(user)を変換する。fileurlが返ってくる。
		macrofolder =  unohelper.fileUrlToSystemPath(fileurl)  # fileurlをシステムパスに変換する。マイマクロフォルダへのパス。	
		location = "user"  # マクロの場所。
	relpath = os.path.relpath(filepath, start=macrofolder)  # マクロフォルダからの相対パスを取得。パス区切りがOS依存で返ってくる。
	return "vnd.sun.star.script:{}${}?language=Python&location={}".format(relpath.replace(os.sep, "|"), "{}", location)  # ScriptingURLのbaseurlを取得。Windowsのためにos.sepでパス区切りを置換。	
g_exportedScripts = macro, #マクロセレクターに限定表示させる関数をタプルで指定。		
if __name__ == "__main__":  # オートメーションで実行するとき
	def automation():  # オートメーションのためにglobalに出すのはこの関数のみにする。
		import officehelper
		from functools import wraps
		import sys
		from com.sun.star.beans import PropertyValue  # Struct
		from com.sun.star.script.provider import XScriptContext  
		def connectOffice(func):  # funcの前後でOffice接続の処理
			@wraps(func)
			def wrapper():  # LibreOfficeをバックグラウンドで起動してコンポーネントテクストとサービスマネジャーを取得する。
				try:
					ctx = officehelper.bootstrap()  # コンポーネントコンテクストの取得。
				except:
					print("Could not establish a connection with a running office.", file=sys.stderr)
					sys.exit()
				print("Connected to a running office ...")
				smgr = ctx.getServiceManager()  # サービスマネジャーの取得。
				print("Using {} {}".format(*_getLOVersion(ctx, smgr)))  # LibreOfficeのバージョンを出力。
				return func(ctx, smgr)  # 引数の関数の実行。
			def _getLOVersion(ctx, smgr):  # LibreOfficeの名前とバージョンを返す。
				cp = smgr.createInstanceWithContext('com.sun.star.configuration.ConfigurationProvider', ctx)
				node = PropertyValue(Name = 'nodepath', Value = 'org.openoffice.Setup/Product' )  # share/registry/main.xcd内のノードパス。
				ca = cp.createInstanceWithArguments('com.sun.star.configuration.ConfigurationAccess', (node,))
				return ca.getPropertyValues(('ooName', 'ooSetupVersion'))  # LibreOfficeの名前とバージョンをタプルで返す。
			return wrapper
		@connectOffice  # createXSCRIPTCONTEXTの引数にctxとsmgrを渡すデコレータ。
		def createXSCRIPTCONTEXT(ctx, smgr):  # XSCRIPTCONTEXTを生成。
			class ScriptContext(unohelper.Base, XScriptContext):
				def __init__(self, ctx):
					self.ctx = ctx
				def getComponentContext(self):
					return self.ctx
				def getDesktop(self):
					return ctx.getByName('/singletons/com.sun.star.frame.theDesktop')  # com.sun.star.frame.Desktopはdeprecatedになっている。
				def getDocument(self):
					return self.getDesktop().getCurrentComponent()
			return ScriptContext(ctx)  
		XSCRIPTCONTEXT = createXSCRIPTCONTEXT()  # XSCRIPTCONTEXTの取得。
		doc = XSCRIPTCONTEXT.getDocument()  # 現在開いているドキュメントを取得。
		doctype = "scalc", "com.sun.star.sheet.SpreadsheetDocument"  # Calcドキュメントを開くとき。
	# 	doctype = "swriter", "com.sun.star.text.TextDocument"  # Writerドキュメントを開くとき。
		if (doc is None) or (not doc.supportsService(doctype[1])):  # ドキュメントが取得できなかった時またはCalcドキュメントではない時
			XSCRIPTCONTEXT.getDesktop().loadComponentFromURL("private:factory/{}".format(doctype[0]), "_blank", 0, ())  # ドキュメントを開く。ここでdocに代入してもドキュメントが開く前にmacro()が呼ばれてしまう。
		flg = True
		while flg:
			doc = XSCRIPTCONTEXT.getDocument()  # 現在開いているドキュメントを取得。
			if doc is not None:
				flg = (not doc.supportsService(doctype[1]))  # ドキュメントタイプが確認できたらwhileを抜ける。
		return XSCRIPTCONTEXT
	XSCRIPTCONTEXT = automation()  # XSCRIPTCONTEXTを取得。	
	macro()  # マクロの実行。