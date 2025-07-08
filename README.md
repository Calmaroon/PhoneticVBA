Replicate Python Abydos library in VBA. 

This project will consist of a number of modules/classes to perform phonetic encoding, stemming, tokenization, and distance.

Code can be used in Excel/Access UDFs.

Code Sample:
```
sub EncodeExamples
  debug.print Soundex("Williams")        'Returns W452
  debug.print RefinedSoundex("Williams") 'Returns W783
  debug.print NYSIIS("Williams")         'Returns WALAN
  debug.print Caverphone("Williams")     'Returns WLMS111111
  debug.print Ainsworth("Williams")      'Returns wɪllɪæms
end sub
```

Phonetic Encoding Status:
|Encoding Name|Status|
|-------------|------|
|American Soundex|Complete|
|Daitch-Mokotoff Soundex|Complete|
|Fuzzy Soundex|Complete|
|Refined Soundex|Complete|
|Metaphone|Complete|
|NYSIIS|Complete|
|Caverphone|Complete|
|Statistics Canada|Complete|
|Match Rating Algorithm (MRA)|Complete|
|LEIN|Complete|
|Koelner (Cologne)|Complete|
|FONEM|Complete|
|Davidson's Consonant Code|Complete|
|Ainsworth|Complete|
|SoundD|Complete|
|ParmarKumbarana|Complete|
|Phonex|Complete|
|Phonix|Complete|
|Oxford Name Compression Algorithm (ONCA)|Complete|
|Phonetic Spanish|Complete|
|PHONIC|Complete|
|Phonem|Complete|
|Eudex|In Progress [Need to Code LargeInt --> String workaround]|
|Robert C. Russell's Index|Not Started|
|Double Metaphone|Not Started|
|SoundexBR|Not Started|
|PSHP Soundex/Viewex Coding|Not Started|
|MetaSoundex|Not Started|
|Norphone|Not Started|
|Roger Root|Not Started|
|Alpha Search Inquiry System|Not Started|
|phonet|Not Started|
|SfinxBis|Not Started|
|Standardized Phonetic Frequency Code|Not Started|
|Haase Phonetik|Not Started|
|Reth-Schek Phonetik|Not Started|
|Henry Early|Not Started|
|Dolby Code|Not Started|
|Spanish Metaphone|Not Started|
|NRL English-to-phoneme|Not Started|
|Beider-Morse Phonetic Matching|Not Started|

String Fingerprint Status:
**|Fingerprinter Name|Status|**


Tokenization Status:
|Tokenizer Name|Status|
|-------------|------|
|Baseline Tokenizer|Not Started|
|Character Tokenizer|Not Started|
|Whitespace Tokenizer|Not Started|
|Word Punctuation Tokenizer|Not Started|
|QGram Tokenizer|Not Started|
|QSkipGram Tokenizer|Not Started|
|CV Cluster Tokenizer|Not Started|
|VC Cluster Tokenizer|Not Started|
|RegExp Tokenizer|Not Started|
|C or V Tokenizer|Not Started|
