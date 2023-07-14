import spacy

def find_english_words(word_list):
    nlp = spacy.load("en_core_web_sm")
    english_words = []

    for word in word_list:
        doc = nlp(word)
        if doc.vocab[word].is_stop or doc.vocab[word].is_punct:
            continue
        english_words.append(word)

    return english_words

# Example usage
words = ["cat", "dog", "house", "table", "घर", "मेज"]
english_words = find_english_words(words)
print(english_words)
