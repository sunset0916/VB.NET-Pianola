# VB.NET-Pianola
WindowsAPIからMIDI音源を呼び出すVB.NETの自動演奏ピアノ
# 対応する楽譜ファイルの作り方
1,C4を0としてそこからどれだけ上下したかを調べて出す音の数字を出します  
2,表計算ソフトを使用して縦方向に音の数字を入力していきます  
　音の長さは何もしないイベントという意味でpを入力します  
　音を止めるイベント（休符）は99を入力します  
　一行ごとの時間は入力する曲の楽譜の最小単位にします  
　例：最小が16分音符→16分音符　最小が付点16分音符→32分音符  
3,それを6列入力します  
　6音までの和音に対応しています  
　6音使わない場合でも休符等ですべて埋めます  
4,完成したファイルはカンマ区切りのCSVファイルで保存します  

# 以下ファイル例
「ハルトマンの妖怪少女」最小単位は16分音符で曲の最初だけです  

99,99,99,99,99,99  
-32,-20,99,4,11,99  
p,p,p,p,p,p  
p,p,p,7,99,p  
p,p,p,p,p,p  
-25,99,p,4,p,p  
p,p,p,p,p,p  
-31,-19,p,5,9,14  
p,p,p,p,p,p  
p,p,p,p,p,p  
-24,99,p,5,9,12  
p,p,p,p,p,p  
p,p,p,p,p,p  
-19,p,p,5,9,14  
p,p,p,p,p,p  
-32,-20,p,4,11,99  
p,p,p,p,p,p  
p,p,p,7,99,p  
p,p,p,p,p,p  
p,p,p,4,p,p  
p,p,p,p,p,p  
-33,-21,p,3,10,p  
p,p,p,p,p,p  
p,p,p,6,99,p  
p,p,p,p,p,p  
p,p,p,3,p,p  
p,p,p,p,p,p  
-33,-21,p,3,10,p  
p,p,p,p,p,p  
