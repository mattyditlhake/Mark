import pandas as pd
import spacy
from scispacy.abbreviation import AbbreviationDetector

nlp = spacy.load("en_core_web_sm")

#Add the abbreviation detector to the pipeline
nlp.add_pipe("abbreviation_detector")

def resolve_abbreviation(text):
    doc = nlp(text)
    mapping = {}
    for abrv in doc._.abbreviations:
# Resolve the abbreviation to its full form
        mapping[str(abrv)]=str(abrv._.long_form)
    replaced_text = text
    for abrv, long_form in mapping.items():
        replaced_text = replaced_text.replace(abrv, long_form)
    return replaced_text

#Read the Excel file that has student Answers
Studentscript = pd.read_excel("StudentScript.xlsx", sheet_name="Sheet1")

#Read the Excel file that has the correct Answers
Memo = pd.read_excel("Memo.xlsx", sheet_name="Sheet1")

#Merge the two dataframes
Compare = pd.merge(Studentscript, Memo, on="QuestionNumber", suffixes=('_Student', '_Memo'))
#Compare['Mark']= Compare.apply(lambda row: 2 if row['Answers_Student'] == row['Answers_Memo'] else 0, axis=1)

#Identify Entities that are similar in both student and memo answers
def contains_named_entity_match(Answer_student, answer_memo):
    
    student_response = nlp(Answer_student)
    memo_response = nlp(answer_memo)
    
    student_entities = {ent.text for ent in student_response.ents}
    memo_entities = {ent.text for ent in memo_response.ents}
    
    return len(student_entities.intersection(memo_entities)) > 0

#Using Spacy(Open source NLP library) to compare student answers with memo answers
def spacy_similarity(Answer_student,Answer_memo,threshold=0.6):
   student_response = nlp(Answer_student)
   memo_response = nlp(Answer_memo)
   similarity = student_response.similarity(memo_response)
   return 2 if similarity >= threshold else 0

#Combine the marking logic
def combined_marking(Answer_student, Answer_memo, threshold=0.6):
    
    resolved_student = resolve_abbreviation(Answer_student)
    resolved_memo = resolve_abbreviation(Answer_memo)
    
    if contains_named_entity_match(Answer_student, Answer_memo):
        return 2
    else:
        return spacy_similarity(Answer_student, Answer_memo, threshold)
    
#For the semantic evaluation using spacy(An open-sourse NLP library)
#The evaluation kicks in if the response has greater than 7 characters
def smart_marking(row):
    if len(str(row['Answers_Student'])) > 7:
        return combined_marking(row['Answers_Student'], row['Answers_Memo'])
    else:
        return 2 if row['Answers_Student'] == row['Answers_Memo'] else 0
    
#Compare['Mark']= Compare.apply(lambda row: 2 if row['Answers_Student'] == row['Answers_Memo'] else 0, axis=1)
Compare['Mark'] = Compare.apply(smart_marking, axis=1)

#Add up the total marks
total_marks = Compare['Mark'].sum()
#Update the Compare DataFrame with the total marks
Compare.loc['Total'] = ['','','',total_marks]

#Save the results to an Excel file
Compare.to_excel("Marking.xlsx", index=False)