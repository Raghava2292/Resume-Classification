# Resume-Classification
An application that will predict the category of a resume among a total of 4 categories.

The data contains the .doc, .docx., and .pdf files of 4 categories of resumes. These categories are the PeopleSoft, React JS, SQL Developer, and the WorkDay Developer.
So, the text data is extracted from the resumes with various tools like textract, win32com, and pypdf. Then text preprocessing is done.

After the NLP techniques are used to analyze and vectorize the data.

Once the vectorization is done, various models are trained on this data to make the classifications.

The random forest model gave the best results. So using the trained random forest model, an application that will take in a resume and will predict the category to which that resume belongs to. A function to extract the skills from the resume is also written. This function extracts the skills from the resume.

