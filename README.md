This collection of VBA modules/classes to perform phonetic encoding, stemming, tokenization, and distance.

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
|Encoding Name|Status|Ported From|
|-------------|------|------------|
|American Soundex|Complete|Abydos|
|Daitch-Mokotoff Soundex|Complete|Abydos|
|Fuzzy Soundex|Complete|Abydos|
|Refined Soundex|Complete|Abydos|
|PSHP Soundex/Viewex Coding|Complete|Abydos|
|SoundexBR|Complete|Abydos|
|Robert C. Russell's Index|Complete|Abydos|
|Roger Root|Complete|Abydos|
|Metaphone|Complete|Abydos|
|Double Metaphone|Complete|Abydos|
|Spanish Metaphone|Complete|Abydos|
|NYSIIS|Complete|Abydos|
|Caverphone|Complete|Abydos|
|Statistics Canada|Complete|Abydos|
|Match Rating Algorithm (MRA)|Complete|Abydos|
|LEIN|Complete|Abydos|
|Koelner (Cologne)|Complete|Abydos|
|Reth-Schek Phonetik|Complete|Abydos|
|FONEM|Complete|Abydos|
|Davidson's Consonant Code|Complete|Abydos|
|Ainsworth|Complete|Abydos|
|NRL English-to-phoneme|Complete|Abydos|
|SoundD|Complete|Abydos|
|ParmarKumbarana|Complete|Abydos|
|Phonex|Complete|Abydos|
|Phonix|Complete|Abydos|
|Oxford Name Compression Algorithm (ONCA)|Complete|Abydos|
|Phonetic Spanish|Complete|Abydos|
|PHONIC|Complete|Abydos|
|Phonem|Complete|Abydos|
|Eudex|In Progress [Need to Code LargeInt --> String workaround]|
|Norphone|In Progress||
|MetaSoundex|Not Started||
|Alpha Search Inquiry System|Not Started||
|phonet|Not Started||
|SfinxBis|Not Started||
|Standardized Phonetic Frequency Code|Not Started||
|Haase Phonetik|Not Started||
|Henry Early|Not Started||
|Dolby Code|Not Started||
|Beider-Morse Phonetic Matching|Not Started||

Stemmer Status:
|Stemmer Name|Status|Ported From|
|-------------|------|------------|
|Porter|Completed|Abydos|
|Porter2|Not Started||

String Fingerprint Status:
|Fingerprinter Name|Status|
|-------------|------|

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
