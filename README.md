Replicate Python Abydos library in VBA. 

This project will consist of a number of modules/classes to perform phonetic encoding, stemming and distance.

Code Sample:
```
sub EncodeExamples
  debug.print Soundex("Williams") 'Returns W452
  debug.print NYSIIS("Williams") 'Returns WALAN
  debug.print Ainsworth("Williams") 'Returns wɪllɪæms
end sub
```

Phonetic Encoding Status:
|Encoding Name|Status|
|-------------|------|
|Caverphone|Complete|
|American Soundex|Complete|
|Daitch-Mokotoff Soundex|Complete|
|Metaphone|Complete|
|Lein|Complete|
|Koelner (Cologne)|Complete|
|NYSIIS|Complete|
|Davidson's Consonant Code|Complete|
|Ainsworth|Complete|
|SoundD|Complete|
|Match Rating Approach (MRA)|Complete|
|ParmarKumbarana|Complete|
|Oxford Name Compression Algorithm (ONCA)|Complete|
|Robert C. Russell's Index|Not Started|
|Refined Soundex|Not Started|
|Daitch-Mokotoff Soundex|Not Started|
|Double Metaphone|Not Started|
|Phonex|Not Started|
|FONEM|Not Started|
|Norphone|Not Started|
|Beider-Morse|Not Started|
|Roger Root|Not Started|
|Refined Soundex|Not Started|
|Fuzzy Soundex|Not Started|
|SoundexBR|Not Started|
|MetaSoundex|Not Started|
|Statistics Canada|Not Started|

