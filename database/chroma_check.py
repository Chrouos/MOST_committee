import chromadb

# 連接到資料庫
client = chromadb.PersistentClient(path="./chroma_database/")

# 列出所有集合名稱
collections = client.list_collections()
print("Collections:", collections)

if collections:
    for collection_name in collections:
        # 使用 get_collection 根據名稱取得集合
        collection = client.get_collection(name=collection_name)
        
        # 列出集合內容
        print(f"Collection: {collection_name}")
        data = collection.get()
        print("Data:", data)
else:
    print("No collections found in the database.")
