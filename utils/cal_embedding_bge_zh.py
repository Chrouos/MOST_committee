from langchain_community.embeddings import HuggingFaceBgeEmbeddings
import numpy as np

model_name = 'BAAI/bge-large-zh-v1.5'
model_kwargs = {'device':'cpu'}
encode_kwargs = {'normalize_embeddings':True}
model = HuggingFaceBgeEmbeddings(
    model_name = model_name,
    model_kwargs = model_kwargs,
    encode_kwargs = encode_kwargs,
    # query_instruction="Represent this sentence for searching relevant passages:"
)

def get_embeddings_zh():
    return model

def calculate_query_embedding(query):
    return model.embed_query(query)

def calculate_docs_embedding_zh(docs):
    return model.embed_documents(docs)

def calculate_cosine_sim(a, b):
    
    # Compute the dot product of a and b
    dot_product = np.dot(a, b)
    
    # Compute the L2 norm of a and b
    norm_a = np.linalg.norm(a)
    norm_b = np.linalg.norm(b)
    
    # Compute the cosine similarity
    sim = dot_product / (norm_a * norm_b)
    return sim
