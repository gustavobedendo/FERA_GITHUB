<<<<<<< HEAD
version https://git-lfs.github.com/spec/v1
oid sha256:19ec237f22306ac682acb5a1a89e2e73674b6024361614c8b56a77d93b6aac4c
size 1091
=======
import sys
import lucene
 
from java.io import File
from org.apache.lucene.analysis.standard import StandardAnalyzer
from org.apache.lucene.document import Document, Field
from org.apache.lucene.search import IndexSearcher
from org.apache.lucene.index import IndexReader
from org.apache.lucene.queryparser.classic import QueryParser
from org.apache.lucene.store import SimpleFSDirectory
from org.apache.lucene.util import Version
 
if __name__ == "__main__":
    lucene.initVM()
    analyzer = StandardAnalyzer(Version.LUCENE_4_10_1)
    reader = IndexReader.open(SimpleFSDirectory(File(r"D:\35584-17\IPED_Eq01\iped\index")))
    searcher = IndexSearcher(reader)
 
    query = QueryParser(Version.LUCENE_4_10_1, "text", analyzer).parse("Find this sentence please")
    MAX = 1000
    hits = searcher.search(query, MAX)
 
    print(f"Found {hits.totalHits} document(s) that matched query '{query}':")
    for hit in hits.scoreDocs:
        print(hit.score, hit.doc, hit.toString())
        doc = searcher.doc(hit.doc)
        print(doc.get("text").encode("utf-8"))
>>>>>>> 9c2d710ecefe0bb84324b1f01b7384e5810496b7
