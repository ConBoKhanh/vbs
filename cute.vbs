Dim spV
Set spV = CreateObject("SAPI.spVoice")
Set spV.Voice = spV.GetVoices.Item(1)
SpV.Speak "HELLO"
Set spV.Voice = spV.GetVoices.Item(0)
SpV.Speak "Hi,How are you"
Set spV.Voice = spV.GetVoices.Item(1)
SpV.Speak "Not bad , and you"
Set spV.Voice = spV.GetVoices.Item(0)
SpV.Speak "Same"
Set spV.Voice = spV.GetVoices.Item(1)
SpV.Speak "wanna hear a funny story"
Set spV.Voice = spV.GetVoices.Item(0)
SpV.Speak "okay"
Set spV.Voice = spV.GetVoices.Item(1)
SpV.Speak "a lot of people say SAKURA SHIODA is cute L O L"
Set spV.Voice = spV.GetVoices.Item(0)
SpV.Speak "but that is the true , she is cute"
Set spV.Voice = spV.GetVoices.Item(1)
SpV.Speak "what the fuck"