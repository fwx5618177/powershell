# Word Guess Game #
It is a PowerShell script to generate PowerPoint files with random words from corpus for the "someone acts, someone guesses the word/sentence" (你比我猜). You can specific the corpus location and output file location and word count during the generation. 

It use the [THUOCL：清华大学开放中文词库](http://thuocl.thunlp.org/) corpus as default data source. 
Sample command:

```.\generate.ps1 -FilePath .\data\THUOCL_chengyu.txt -Size 50 -OutputFile test.pptx ```
