import pandas as pd
from newspaper import Article, ArticleException
import os
from nltk.tokenize import word_tokenize, sent_tokenize
from nltk.corpus import cmudict
import re


#counting personal pronouns
def count_personal_pronouns(article):
    pronoun_list = ['I', 'me', 'my', 'mine', 'myself',
                    'you', 'your', 'yours', 'yourself',
                    'he', 'him', 'his', 'himself',
                    'she', 'her', 'hers', 'herself',
                    'it', 'its', 'itself',
                    'we', 'us', 'our', 'ours', 'ourselves',
                    'they', 'them', 'their', 'theirs', 'themselves']
    pronoun_regex = r'\b(?:{})\b'.format('|'.join(pronoun_list))
    matches = re.findall(pronoun_regex, article, flags=re.IGNORECASE)
    pronoun_count = len(matches)

    # Exclude "US" from the count
    us_count = article.count("US")
    pronoun_count -= us_count

    return pronoun_count





#counting no of syllable
def count_syllables(word):
    exceptions = ['es','ed']

    if any(word.endswith(exception) for exception in exceptions):
        return 0
    
    vowels = 'aeiouAEIOU'
    count = 0
    prev_char = None

    for char in word:
        if char in vowels:
            if prev_char is None or prev_char not in vowels:
                count += 1
        prev_char = char
    
    return count


#importing the excel file
data = pd.read_excel('Input.xlsx')

#reading stop word files
stop_word = []
files = os.listdir('StopWords/')
for i in files:
    with open('StopWords/'+i,'r',encoding='latin-1') as file:
        for line in file:
            line.rstrip()
            stop_word.append(line.split()[0])

#readin positive words
positive_words = []
with open('MasterDictionary/positive-words.txt','r',encoding='latin-1') as file:
        for line in file:
            line.rstrip()
            positive_words.append(line.split()[0])

#readin positive words
negative_words = []
with open('MasterDictionary/negative-words.txt','r',encoding='latin-1') as file:
        for line in file:
            line.rstrip()
            negative_words.append(line.split()[0])


data['POSITIVE SCORE'] = 0
data['NEGATIVE SCORE'] = 0
data['POLARITY SCORE'] = 0
data['SUBJECTIVITY SCORE'] = 0
data['AVG SENTENCE LENGTH'] = 0
data['PERCENTAGE OF COMPLEX WORDS'] = 0
data['FOG INDEX'] = 0
data['AVG NUMBER OF WORDS PER SENTENCE'] = 0
data['COMPLEX WORD COUNT'] = 0
data['WORD COUNT'] = 0
data['SYLLABLE PER WORD'] = 0
data['PERSONAL PRONOUNS'] = 0
data['AVG WORD LENGTH'] = 0


for i in range(len(data)):
    
    #web scrapping
    ar = data['URL'][i]
    try:
        article = Article(ar)
        article.download()
        article.parse()
        article_text = article.text


        #tokenizing
        word_tokenize_article = word_tokenize(article_text)
        sentance_tokenize_article = sent_tokenize(article_text)

        #cleaning of the article
        #removing punctuation marks and special characters
        new_article = []
        pattern = re.compile('[^a-zA-Z]')
        for word in word_tokenize_article:
            word = re.sub(pattern, '', word)
            if word:
                new_article.append(word)
        
        #removing stop words
        cleaned_aricle = []
        for word in new_article:
            if word.upper() not in stop_word:
                cleaned_aricle.append(word)

        #calculating positive and negative score
        positive_score = 0
        negative_score = 0
        for word in cleaned_aricle:
            if word in positive_words:
                positive_score += 1
            elif word in negative_words:
                negative_score += 1

        #polarity score
        polarity_score = (positive_score - negative_score) / ((positive_score + negative_score) + 0.000001)
        
        #subjectivity score
        subjectivity_score = (positive_score + negative_score) / ((len(cleaned_aricle)) + 0.000001)

        word_count = len(cleaned_aricle)
        sentance_count = len(sentance_tokenize_article)

        #Analysis of readability
        #average sentance lenght 
        average_sentance_lenght = len(new_article) / sentance_count

        #percentage of complex word
        complex_word_count = 0
        total_syllable_count = 0
        total_word_lenght = 0
        for word in new_article:
            total_word_lenght += len(word)
            temp = count_syllables(word)
            total_syllable_count += temp
            if temp > 2 :
                complex_word_count += 1
        percentageOfComplexWord = complex_word_count / len(new_article)

        #fog index
        fog_index = 0.4 * (average_sentance_lenght + percentageOfComplexWord)

        #syllable count per word
        syllable_per_word = total_syllable_count / len(new_article)

        #Average word lenght
        average_word_lenght = total_word_lenght / len(new_article)

        #personal pronouns
        personal_pronouns = 0
        for sentance in sentance_tokenize_article:
            personal_pronouns += count_personal_pronouns(sentance)

        data['POSITIVE SCORE'][i] = positive_score
        data['NEGATIVE SCORE'][i] = negative_score
        data['POLARITY SCORE'][i] = polarity_score
        data['SUBJECTIVITY SCORE'][i] = subjectivity_score
        data['AVG SENTENCE LENGTH'][i] = average_sentance_lenght
        data['PERCENTAGE OF COMPLEX WORDS'][i] = percentageOfComplexWord
        data['FOG INDEX'][i] = fog_index
        data['AVG NUMBER OF WORDS PER SENTENCE'][i] = average_sentance_lenght
        data['COMPLEX WORD COUNT'][i] = complex_word_count
        data['WORD COUNT'][i] = word_count
        data['SYLLABLE PER WORD'][i] = syllable_per_word
        data['PERSONAL PRONOUNS'][i] = personal_pronouns
        data['AVG WORD LENGTH'][i] = average_word_lenght
    
    except ArticleException as e:
        print("Download failed with URL: ",ar)
        print("ERROR: ", str(e))

    print("Pass: ",i) 


data.to_excel('output.xlsx')