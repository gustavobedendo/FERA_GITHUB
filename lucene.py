import sys
import os
import utilities_general
from java.lang import System
from java.net import URL, URLClassLoader

# Paths to the Jython and Lucene JAR files
jython_jar = os.path.join(utilities_general.get_application_path(), "Lucene", "jython-standalone-2.7.4.jar")
lucene_core_jar = os.path.join(utilities_general.get_application_path(), "Lucene", "lucene-core-6.0.0.jar")
lucene_analyzers_jar = os.path.join(utilities_general.get_application_path(), "Lucene", "lucene-analyzers-common-6.0.0.jar")

global added_paths
# List to store added paths for later removal
added_paths = []

# Function to add to sys.path if the path is not already included
def add_to_sys_path(path):
    global added_paths
    if path not in sys.path:
        sys.path.append(path)
        added_paths.append(path)  # Track the path for later removal

# Function to add JARs to the classpath for Jython
def add_lucene_to_classpath():
    global added_paths
    # Add JARs to the classpath using URLClassLoader
    add_to_sys_path(jython_jar)
    add_to_sys_path(lucene_core_jar)
    add_to_sys_path(lucene_analyzers_jar)
    
    # Create a URLClassLoader to load JAR files into the classpath
    urls = [
        URL("file://" + jython_jar),
        URL("file://" + lucene_core_jar),
        URL("file://" + lucene_analyzers_jar),
    ]
    
    class_loader = URLClassLoader(urls, System.getClassLoader())
    System.setProperty("java.class.path", System.getProperty("java.class.path") + ":" + ":".join([str(url) for url in urls]))

# Function to remove JAR paths from sys.path after execution
def remove_lucene_from_path():
    global added_paths
    # Remove added paths after execution
    for path in added_paths:
        if path in sys.path:
            sys.path.remove(path)

# Add Lucene JARs to the classpath
add_lucene_to_classpath()

# Now you should be able to import Lucene classes
try:
    from org.apache.lucene.analysis.standard import StandardAnalyzer
    from org.apache.lucene.index import DirectoryReader
    from org.apache.lucene.queryparser.classic import QueryParser
    from org.apache.lucene.search import IndexSearcher
    from org.apache.lucene.store import FSDirectory
    import java.nio.file.Paths

    # Path to the existing index directory (replace with your index path)
    index_path = r"D:\28875-23-Anexo\IPED\iped\index"

    # Open the index directory
    directory = FSDirectory.open(Paths.get(index_path))

    # Open the DirectoryReader to read the index
    reader = DirectoryReader.open(directory)

    # Set to store unique field names
    field_names_set = set()

    # Iterate over all the documents in the index and get the field names
    for doc_id in range(reader.maxDoc()):
        if reader.isDeleted(doc_id):
            continue  # Skip deleted documents

        doc = reader.document(doc_id)
        # Add field names to the set (set will handle duplicates automatically)
        for field in doc.getFields():
            field_names_set.add(field.name())

    # Print unique field names
    for field_name in field_names_set:
        print(f"Field name: {field_name}")

    # Close resources
    reader.close()
    directory.close()

finally:
    # Ensure cleanup of classpath after execution
    remove_lucene_from_path()
