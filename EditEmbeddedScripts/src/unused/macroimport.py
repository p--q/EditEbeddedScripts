#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
global XSCRIPTCONTEXT
def macro(documentevent=None):  # 引数は文書のイベント駆動用。  
	doc = XSCRIPTCONTEXT.getDocument() if documentevent is None else documentevent.Source  # ドキュメントのモデルを取得。 
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
	simplefileaccess = smgr.createInstanceWithContext("com.sun.star.ucb.SimpleFileAccess", ctx)  # SimpleFileAccess
	
	controller = doc.getCurrentController()  # コントローラの取得。
# 	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	
# 	import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)  # デバッグできない。
	install_meta, dummy = tdocimport(simplefileaccess)
	install_meta(getModuleFolderPath(ctx, smgr, doc))
	
# 	install_meta('http://localhost:15000')
	
	import consts
	s = consts.LISTSHEET["name"]
	
	sheet = controller.getActiveSheet()
	sheet["A1"].setString(s)
	
	
	
# 	import fib
# 	import spam
# 	import grok.blah
# 	print(grok.blah.__file__)
def getModuleFolderPath(ctx, smgr, doc):
	transientdocumentsdocumentcontentfactory = smgr.createInstanceWithContext("com.sun.star.frame.TransientDocumentsDocumentContentFactory", ctx)
	transientdocumentsdocumentcontent = transientdocumentsdocumentcontentfactory.createDocumentContent(doc)
	tdocurl = transientdocumentsdocumentcontent.getIdentifier().getContentIdentifier()  # ex. vnd.sun.star.tdoc:/1	
	return "/".join((tdocurl, "Scripts/python/pythonpath"))  # 開いているドキュメント内の埋め込みマクロフォルダへのパス。	
def tdocimport(simplefileaccess):
	import sys
	import importlib.abc
	from types import ModuleType
# 	from urllib.request import urlopen
# 	from urllib.error import HTTPError, URLError
# 	from html.parser import HTMLParser
	
	
	# Get links from a given URL
	def _get_links(url):
		return simplefileaccess.getFolderContents(url, True)
		
		
		
# 		class LinkParser(HTMLParser):
# 			def handle_starttag(self, tag, attrs):
# 				if tag == 'a':
# 					attrs = dict(attrs)
# 					links.add(attrs.get('href').rstrip('/'))
# 	
# 		links = set()
# 		try:
# 			u = urlopen(url)
# 			parser = LinkParser()
# 			parser.feed(u.read().decode('utf-8'))
# 		except Exception:
# 			pass
# 		return links
	
	class UrlMetaFinder(importlib.abc.MetaPathFinder):
		def __init__(self, baseurl):
			self._baseurl = baseurl
			self._links   = {}
			self._loaders = {baseurl: UrlModuleLoader(baseurl)}
		def find_module(self, fullname, path=None):
			
# 			import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)  
			
			
			if path is None:
				baseurl = self._baseurl
			else:
				if not path[0].startswith(self._baseurl):
					return None
				baseurl = path[0]
			parts = fullname.split('.')
			basename = parts[-1]
			if basename not in self._links:  # Check link cache
				self._links[baseurl] = _get_links(baseurl)
			if basename in self._links[baseurl]:  # Check if it's a package
				fullurl = "/".join((self._baseurl, basename))
				loader = UrlPackageLoader(fullurl)
				try:  # Attempt to load the package (which accesses __init__.py)
					loader.load_module(fullname)
					self._links[fullurl] = _get_links(fullurl)
					self._loaders[fullurl] = UrlModuleLoader(fullurl)
				except ImportError:
					loader = None
				return loader
			filename = "".join((basename, '.py'))
			if filename in self._links[baseurl]:  # A normal module
				return self._loaders[baseurl]
			else:
				return None
		def invalidate_caches(self):
			self._links.clear()
	class UrlModuleLoader(importlib.abc.SourceLoader):  # Module Loader for a URL
		def __init__(self, baseurl):
			self._baseurl = baseurl
			self._source_cache = {}
		def module_repr(self, module):
			return '<urlmodule {} from {}>'.format(module.__name__, module.__file__)
		def load_module(self, fullname):  # Required method
			code = self.get_code(fullname)
			mod = sys.modules.setdefault(fullname, ModuleType(fullname))
			mod.__file__ = self.get_filename(fullname)
			mod.__loader__ = self
			mod.__package__ = fullname.rpartition('.')[0]
			exec(code, mod.__dict__)
			return mod
		def get_code(self, fullname):  # Optional extensions
			src = self.get_source(fullname)
			return compile(src, self.get_filename(fullname), 'exec')
		def get_data(self, path):
			pass
		def get_filename(self, fullname):
			return "".join((self._baseurl, '/', fullname.split('.')[-1], '.py'))
		def get_source(self, fullname):
			filename = self.get_filename(fullname)
			if filename in self._source_cache:
				return self._source_cache[filename]
			try:
				
				inputstream = simplefileaccess.openFileRead(fullname)
				dummy, b = inputstream.readBytes([], inputstream.available())  # simplefileaccess.getSize(module_tdocurl)は0が返る。
				source = bytes(b).decode("utf-8")  # モジュールのソースをテキストで取得。
				
				
# 				u = urlopen(filename)
# 				source = u.read().decode('utf-8')
				
				
				
				
				
				self._source_cache[filename] = source
				return source
			except:
				raise ImportError("Can't load {}".format(filename))
		def is_package(self, fullname):
			return False
	class UrlPackageLoader(UrlModuleLoader):  # Package loader for a URL
		def load_module(self, fullname):
			mod = super().load_module(fullname)
			mod.__path__ = [ self._baseurl ]
			mod.__package__ = fullname
		def get_filename(self, fullname):
			return "/".join((self._baseurl, '__init__.py'))
		def is_package(self, fullname):
			return True
	_installed_meta_cache = { }
	def install_meta(address):  # Utility functions for installing the loader
		if address not in _installed_meta_cache:
			finder = UrlMetaFinder(address)
			_installed_meta_cache[address] = finder
			sys.meta_path.append(finder)
	def remove_meta(address):  # Utility functions for uninstalling the loader
		if address in _installed_meta_cache:
			finder = _installed_meta_cache.pop(address)
			sys.meta_path.remove(finder)
	return install_meta, remove_meta
g_exportedScripts = macro, #マクロセレクターに限定表示させる関数をタプルで指定。	