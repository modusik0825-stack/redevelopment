# 재개발 위클리 — '대화 raw' 폴더의 카톡 내보내기(.txt / .zip)를 읽어
# 가장 최근 한 주(월~일)를 분석하고, 페이지가 읽는 weekly.js 를 만든다.
#   실행: 업데이트.bat 더블클릭 (또는 powershell -File update.ps1 [-noopen])
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Web
Add-Type -AssemblyName System.IO.Compression.FileSystem
$clock = [Diagnostics.Stopwatch]::StartNew()

$root   = $PSScriptRoot
$rawDir = Get-ChildItem -LiteralPath $root -Directory | Where-Object { $_.Name -like '*raw*' } | Select-Object -First 1
if (-not $rawDir) { Write-Host "'대화 raw' 폴더를 찾지 못했어요."; exit 1 }

# ───────────────────────── 1. 파일 읽기 ─────────────────────────
$dateRe = [regex]'(?m)^(?:(\d{4})\. (\d{1,2})\. (\d{1,2})\. |-+ (\d{4})년 (\d{1,2})월 (\d{1,2})일)'
function Get-LastDate([string]$t) {
  $tail = if ($t.Length -gt 300000) { $t.Substring($t.Length - 300000) } else { $t }
  $ms = $dateRe.Matches($tail)
  if ($ms.Count -eq 0) { return $null }
  $m = $ms[$ms.Count - 1]
  if ($m.Groups[1].Success) { return [datetime]::new([int]$m.Groups[1].Value, [int]$m.Groups[2].Value, [int]$m.Groups[3].Value) }
  return [datetime]::new([int]$m.Groups[4].Value, [int]$m.Groups[5].Value, [int]$m.Groups[6].Value)
}
$texts = New-Object System.Collections.Generic.List[object]
foreach ($f in Get-ChildItem -LiteralPath $rawDir.FullName -File) {
  if ($f.Extension -ieq '.txt') {
    $t = [IO.File]::ReadAllText($f.FullName, [Text.Encoding]::UTF8)
    $texts.Add([pscustomobject]@{ name = $f.Name; text = $t; last = (Get-LastDate $t) })
  } elseif ($f.Extension -ieq '.zip') {
    $zip = [IO.Compression.ZipFile]::OpenRead($f.FullName)
    try {
      foreach ($e in $zip.Entries) {
        if ($e.FullName -notlike '*.txt') { continue }
        $sr = New-Object IO.StreamReader($e.Open(), [Text.Encoding]::UTF8)
        try { $t = $sr.ReadToEnd() } finally { $sr.Dispose() }
        $texts.Add([pscustomobject]@{ name = $f.Name + '::' + $e.FullName; text = $t; last = (Get-LastDate $t) })
      }
    } finally { $zip.Dispose() }
  }
}
if ($texts.Count -eq 0) { Write-Host "'대화 raw' 폴더에 카톡 파일이 없어요."; exit 1 }
$latestDay = ($texts | Where-Object { $_.last } | Measure-Object -Property last -Maximum).Maximum
$dow = (([int]$latestDay.DayOfWeek) + 6) % 7
$wS = $latestDay.Date.AddDays(-$dow)        # 이번 주 월요일
$wE = $wS.AddDays(7)
$pS = $wS.AddDays(-7)                        # 지난 주
$cutoff = $wS.AddDays(-35)
Write-Host ("기준 주: {0:yyyy-MM-dd} ~ {1:yyyy-MM-dd}" -f $wS, $wE.AddDays(-1))

# ───────────────────────── 2. 대화 파싱 + 중복 제거 ─────────────────────────
class Msg {
  [datetime]$t; [string]$a; [string]$n; [string]$x; [int]$s; [bool]$rt
  [string[]]$z; [bool]$pos; [bool]$neg; [bool]$negev; [bool]$q
}
$reMob     = [regex]'^(\d{4})\. (\d{1,2})\. (\d{1,2})\. (오전|오후) (\d{1,2}):(\d{2})(?:, (.+?) :(?: (.*))?|: (.*))$'
$reMobDate = [regex]'^\d{4}년 \d{1,2}월 \d{1,2}일 \S요일$'
$rePcDate  = [regex]'^-+ (\d{4})년 (\d{1,2})월 (\d{1,2})일 \S+ -+$'
$rePcMsg   = [regex]'^\[(.+?)\] \[(오전|오후) (\d{1,2}):(\d{2})\] (.*)$'
$reSys     = [regex]'님이 (들어왔습니다|나갔습니다)\.$|님을 내보냈습니다\.$|님을 초대했습니다\.$|채팅방 관리자가'
function Get-Hour([string]$ap, [string]$h) { $v = [int]$h; if ($ap -eq '오전') { if ($v -eq 12) { 0 } else { $v } } else { if ($v -eq 12) { 12 } else { $v + 12 } } }

$best = @{}; $all = New-Object System.Collections.Generic.List[Msg]; $seq = 0
foreach ($src in $texts) {
  if ($src.last -and $src.last -lt $cutoff) { continue }
  $list = New-Object System.Collections.Generic.List[Msg]
  $isPc = [regex]::IsMatch($src.text.Substring(0, [Math]::Min(200000, $src.text.Length)), '(?m)^\[.+?\] \[(오전|오후) \d{1,2}:\d{2}\] ')
  $cur = $null; $day = $null
  foreach ($raw in $src.text.Split("`n")) {
    $line = $raw.TrimEnd("`r")
    if ($isPc) {
      $m = $rePcDate.Match($line)
      if ($m.Success) { $day = [datetime]::new([int]$m.Groups[1].Value, [int]$m.Groups[2].Value, [int]$m.Groups[3].Value); $cur = $null; continue }
      if ($day) {
        $m = $rePcMsg.Match($line)
        if ($m.Success) {
          $t = $day.AddHours((Get-Hour $m.Groups[2].Value $m.Groups[3].Value)).AddMinutes([int]$m.Groups[4].Value)
          if ($t -lt $cutoff) { $cur = $null; continue }
          $cur = [Msg]::new(); $cur.t = $t; $cur.a = $m.Groups[1].Value.Trim(); $cur.x = $m.Groups[5].Value
          $list.Add($cur); continue
        }
      }
    } else {
      $m = $reMob.Match($line)
      if ($m.Success) {
        if (-not $m.Groups[7].Success) { $cur = $null; continue }
        $t = [datetime]::new([int]$m.Groups[1].Value, [int]$m.Groups[2].Value, [int]$m.Groups[3].Value, (Get-Hour $m.Groups[4].Value $m.Groups[5].Value), [int]$m.Groups[6].Value, 0)
        if ($t -lt $cutoff) { $cur = $null; continue }
        $cur = [Msg]::new(); $cur.t = $t; $cur.a = $m.Groups[7].Value.Trim(); $cur.x = $m.Groups[8].Value
        $list.Add($cur); continue
      }
      if ($reMobDate.IsMatch($line)) { $cur = $null; continue }
    }
    if ($reSys.IsMatch($line)) { $cur = $null; continue }
    if ($cur) { $cur.x += "`n" + $line }
  }
  $cnt = @{}
  foreach ($m in $list) {
    $m.x = $m.x.Replace('메시지가 삭제되었습니다.', '').Trim()
    $key = $m.t.ToString('yyyyMMddHHmm') + '|' + $m.a + '|' + $m.x
    $c = 1 + [int]$cnt[$key]; $cnt[$key] = $c
    if ($c -gt [int]$best[$key]) { $best[$key] = $c; $m.s = $seq++; $all.Add($m) }
  }
}
$ALLM = @($all | Where-Object { $_.a -ne '오픈채팅봇' } | Sort-Object t, s)
foreach ($m in $ALLM) { $m.n = $m.a.Split('/')[0].Trim(); $m.rt = $m.n.Replace(' ', '').StartsWith('리트리') }
Write-Host ("메시지 {0:N0}개 읽음 ({1:N1}초)" -f $ALLM.Count, $clock.Elapsed.TotalSeconds)
function Get-LB([datetime]$d) { $lo = 0; $hi = $ALLM.Count; while ($lo -lt $hi) { $mid = [int][Math]::Floor(($lo + $hi) / 2); if ($ALLM[$mid].t -lt $d) { $lo = $mid + 1 } else { $hi = $mid } }; $lo }

# ───────────────────────── 3. 구역 추출 ─────────────────────────
$reZone = [regex]'(?<![가-힣A-Za-z0-9])([가-힣]{2,5}?)\s?(?:제\s?)?(\d{1,3}(?:-\d{1,2})?)(\s?(?:구역|지구))?(?![억만천원평%년월일시분초층동호개명세대차번위등기살주배회건채선단구박급프인학종룸킬턴타순권편곳군점표티장]|[.,~/]?\d|가구|퍼)'
$stop = New-Object 'System.Collections.Generic.HashSet[string]'
foreach ($w in ('강남 강북 강동 강서 서울 경기 인천 부산 대구 광주 대전 울산 세종 제주 수도권 지방 시즌 비거주 실거주 일시적 압도적 전세 월세 반전세 뉴스 전과 분의 가구 신통 노뉴 상위 하위 최소 최대 지하 지상 ' +
  '무주택 다주택 유주택 주택 오늘 내일 어제 이번 다음 지난 올해 작년 내년 거의 대략 아파트 빌라 토허 국평 전용 공급 분양 입주 조합원 투자 대출 금리 기준 단계 레벨 버전 상품 평형 타입 ' +
  '감평 비례율 프리미엄 갭투자 아이폰 갤럭시 플러스 버스 마을버스 광역 경부 호선 지하철 초품아 신축 구축 재건축 재개발 모아 모아타운 모아주택 정비구역 구역 지구 사업 일반 ' +
  '최근 요즘 그냥 진짜 정말 혹시 아마 이거 저거 그거 여기 저기 거기 이게 저게 그게 제가 저는 우리 이제 아직 벌써 다시 계속 보통 대부분 일단 무조건 결국 하나 둘이 ' +
  '평균 매매 매수 매도 호가 시세 실거래 급매 갭 전용면적 대지 지분 감정 권리 분담금 추가 분담 정도 이상 이하 미만 초과 약간 조금 많이 거래 계약 잔금 중도금 계약금 ' +
  '사진 동영상 파일 이모티콘 샵검색 나이 단독 다가구 다세대 동의 찬성 반대 연식 접도 동의서 등급 점수 순위 사이즈 투룸 원룸 쓰리룸 근생 상가 토지 연봉 월급 소득 자산 현금 대지지분 ' +
  '금전 건강 애정 운세 행운 오전 오후 새벽 아침 저녁 ' +
  '노원구 도봉구 강북구 성북구 동대문구 중랑구 광진구 성동구 용산구 마포구 서대문구 은평구 종로구 중구 영등포구 동작구 관악구 구로구 금천구 양천구 강서구 송파구 강동구 서초구 강남구 ' +
  '레미안 래미안 자이 힐스 푸르지오 아이파크 롯데 현대 삼성 대우 디에이치 아크로 오티에르 써밋 르엘 위브 센트럴 리센츠 엘스 트리지움 파크리오 헬리오').Split(' ')) { if ($w) { [void]$stop.Add($w) } }
$reStrip = [regex]'https?://\S+|@[^\s@]+(?:\s?/\s?[^\s@/]+)*'
$suffixed = New-Object 'System.Collections.Generic.HashSet[string]'
$rawZ = @{}
for ($i = 0; $i -lt $ALLM.Count; $i++) {
  $s = $reStrip.Replace($ALLM[$i].x, ' ')
  $found = New-Object System.Collections.Generic.List[string]
  foreach ($mm in $reZone.Matches($s)) {
    $p = $mm.Groups[1].Value
    if ($p.Length -ge 3 -and $p.EndsWith('동')) { $p = $p.Substring(0, $p.Length - 1) }
    if ($stop.Contains($p) -or $p -match '(율|률|노후도|비용|공사비)$') { continue }
    $k = $p + $mm.Groups[2].Value
    if ($mm.Groups[3].Success) { [void]$suffixed.Add($k) }
    if (-not $found.Contains($k)) { $found.Add($k) }
  }
  if ($found.Count) { $rawZ[$i] = $found }
}
$zc = @{}; foreach ($v in $rawZ.Values) { foreach ($k in $v) { $zc[$k] = 1 + [int]$zc[$k] } }
$authorsStr = (($ALLM | ForEach-Object { $_.a.Replace(' ', '').Replace('/', '') } | Sort-Object -Unique) -join "`n")
$ok = New-Object 'System.Collections.Generic.HashSet[string]'
foreach ($k in $zc.Keys) { if ($zc[$k] -ge 3 -and ($suffixed.Contains($k) -or $zc[$k] -ge 100 -or -not $authorsStr.Contains($k))) { [void]$ok.Add($k) } }

# ───────────────────────── 4. 추천·부정·질문 신호 ─────────────────────────
$POS   = [regex]'(?i)추천|강추|한\s?표|원픽|좋게\s?봅|좋아\s?보|괜찮아\s?보|괜찮(아요|습니다|죠|은\s?곳)|사세요|사두|사놓으면|담으세요|지켜보시길|관심\s?가져|저평가|가성비|안전마진|싸(요|다|네|게)|제2의|급이라|사기캐|갓|goat|킹|1순위|저라면|저라도|저도\s?[가-힣0-9]+(요|이요)$'
$NEG   = [regex]'비추|별로|비싸|고점|거르|피하|위험|애매|조심|노답|마피|손절|매도\s?하|사지\s?마|안\s?사'
$NEGEV = [regex]'취소|해제|무산|반려|탈락|빠져|빠지|빠꾸|미끄러|좌초|중단|철회|부결|안\s?나오|안\s?되|못\s?하|ㅠ'
$Q     = [regex]'\?|까요|나요|어때|어떤가|어떨|인가요|할까|될까|맞나'
$TOPIC = [regex]'피\s?\d|[pP]\s?\d|프리미엄|매물|호가|분담금|감평|감정|권리가액|비례율|관처|관리처분|사시|사업시행|조합|입주권|현청|무허가|실거주|초투|총투|갭|평형|일분|분양|추천|괜찮|좋아|좋은|별로|리스크|입지|역세권|가격|시세|\d억|대지지분|지분|동의율|추진위|시공사|이주|착공'
$NOISE = [regex]'^(사진|사진 \d+장|동영상|이모티콘|삭제된 메시지입니다\.|ㅋ+|ㅎ+|ㄷ+|ㅠ+|\?+)$'
for ($i = 0; $i -lt $ALLM.Count; $i++) {
  $m = $ALLM[$i]
  $m.z = if ($rawZ.ContainsKey($i)) { @($rawZ[$i] | Where-Object { $ok.Contains($_) }) } else { @() }
  $m.pos = $POS.IsMatch($m.x); $m.neg = $NEG.IsMatch($m.x); $m.negev = $NEGEV.IsMatch($m.x); $m.q = $Q.IsMatch($m.x)
}

# ───────────────────────── 5. 도움 함수 ─────────────────────────
function Get-Snip([string]$s, [int]$n) { $s = ($s -replace '\s+', ' ').Trim(); if ($s.Length -gt $n) { $s.Substring(0, $n) + '…' } else { $s } }
function Get-Row([int]$i, [string]$type) { $m = $ALLM[$i]; $x = $m.x; if ($x.Length -gt 400) { $x = $x.Substring(0, 400) + '…' }; ,@($m.n, $m.t.ToString('MM/dd HH:mm'), $x, $type) }
function Get-WeekLabel([datetime]$monday) { $th = $monday.AddDays(3); '{0}월 {1}주차' -f $th.Month, [Math]::Ceiling($th.Day / 7) }

# 구역별 집계 (기간 a~b): 언급 수, 부정 소식 수, 추천한 사람
function Get-ZoneStats([datetime]$a, [datetime]$b) {
  $z = @{}
  for ($i = (Get-LB $a); $i -lt $ALLM.Count -and $ALLM[$i].t -lt $b; $i++) {
    $m = $ALLM[$i]
    foreach ($k in $m.z) {
      if (-not $z.ContainsKey($k)) { $z[$k] = [pscustomobject]@{ z = $k; men = 0; neg = 0; people = @{}; ev = (New-Object System.Collections.Generic.List[int]) } }
      $r = $z[$k]; $r.men++
      if ($m.negev) { $r.neg++ }
      if ($m.pos -and -not $m.neg -and -not $m.negev -and -not $m.q) { $r.people[$m.n] = 1; $r.ev.Add($i) }
    }
  }
  $z
}
function Get-Ranking($stats) {
  $negKeys = @($stats.Values | Where-Object { $_.men -ge 5 -and $_.neg -ge 3 -and ($_.neg / $_.men) -ge 0.25 } | ForEach-Object { $_.z })
  $rank = @($stats.Values | Where-Object { $_.people.Count -gt 0 -and $negKeys -notcontains $_.z } |
    Sort-Object @{ e = { $_.people.Count }; Descending = $true }, @{ e = { $_.men }; Descending = $true }, @{ e = { $_.z }; Descending = $false })
  @{ rank = $rank; neg = $negKeys }
}

# 관련 대화: 구역이 나온 메시지 + 바로 앞 질문 + 이어진 투자 관련 대화(3분, 5개, 다른 구역으로 넘어가면 멈춤)
$weekPlaces = @{}
for ($i = (Get-LB $wS); $i -lt $ALLM.Count; $i++) { foreach ($k in $ALLM[$i].z) { $p = [regex]::Match($k, '^(.+?)\d').Groups[1].Value; if ($p) { $weekPlaces[$p] = 1 + [int]$weekPlaces[$p] } } }
function Get-Chats([datetime]$a, [datetime]$b, [string]$zone, [regex]$re, [regex]$excl) {
  $myPref = if ($zone) { [regex]::Match($zone, '^(.+?)\d').Groups[1].Value } else { '장미' }
  $others = @($weekPlaces.Keys | Where-Object { $weekPlaces[$_] -ge 3 -and $_ -ne $myPref -and -not $myPref.StartsWith($_) -and -not $_.StartsWith($myPref) } | ForEach-Object { [regex]::Escape($_) })
  $otherPlace = if ($others.Count) { [regex]('(?<![가-힣])(' + ($others -join '|') + ')') } else { $null }
  $hit = { param($mm) if ($zone) { $mm.z -contains $zone } else { $re.IsMatch($mm.x) } }
  $pick = New-Object 'System.Collections.Generic.SortedDictionary[int,string]'
  for ($i = (Get-LB $a); $i -lt $ALLM.Count -and $ALLM[$i].t -lt $b; $i++) {
    $m = $ALLM[$i]
    if (-not (& $hit $m)) { continue }
    if ($excl) {
      if ($excl.IsMatch($m.x)) { continue }
      $near = (($i - 20)..($i + 20) | Where-Object { $_ -ge 0 -and $_ -lt $ALLM.Count -and $_ -ne $i -and [Math]::Abs(($ALLM[$_].t - $m.t).TotalMinutes) -le 4 } | ForEach-Object { $ALLM[$_].x }) -join ' '
      if ($excl.IsMatch($near) -and $m.x -notmatch '신길|도림|영등포|장미마을') { continue }
    }
    $pick[$i] = 'k'
    for ($j = $i - 1; $j -ge 0 -and $j -ge $i - 6; $j--) {
      if (($m.t - $ALLM[$j].t).TotalMinutes -gt 3) { break }
      if ($ALLM[$j].a -ne $m.a -and $ALLM[$j].q -and $TOPIC.IsMatch($ALLM[$j].x)) { if (-not $pick.ContainsKey($j)) { $pick[$j] = 'c' }; break }
    }
    $n = 0
    for ($j = $i + 1; $j -lt $ALLM.Count -and $n -lt 5; $j++) {
      $mj = $ALLM[$j]
      if (($mj.t - $m.t).TotalMinutes -gt 3) { break }
      if (& $hit $mj) { break }
      if ($zone) { if (@($mj.z | Where-Object { $_ -ne $zone }).Count) { break } } elseif ($mj.z.Count) { break }
      if ($otherPlace -and $otherPlace.IsMatch($mj.x)) { break }
      $tx = $mj.x.Trim()
      if ($NOISE.IsMatch($tx) -or $tx.Length -lt 6 -or -not $TOPIC.IsMatch($tx)) { continue }
      if (-not $pick.ContainsKey($j)) { $pick[$j] = 'c' }
      $n++
    }
  }
  $rows = New-Object System.Collections.Generic.List[object]
  foreach ($kv in $pick.GetEnumerator()) { $rows.Add((Get-Row $kv.Key $kv.Value)) }
  ,$rows
}
function Get-People($rows) { @($rows | ForEach-Object { $_[0] } | Sort-Object -Unique).Count }
function Get-KwCount($rows) { @($rows | Where-Object { $_[3] -eq 'k' }).Count }

# ───────────────────────── 6. 주요 추천 구역 ─────────────────────────
$statW = Get-ZoneStats $wS $wE
$statP = Get-ZoneStats $pS $wS
$rkW = Get-Ranking $statW
$rkP = Get-Ranking $statP
$prevPos = @{}; for ($r = 0; $r -lt $rkP.rank.Count; $r++) { $prevPos[$rkP.rank[$r].z] = $r + 1 }
$picks = New-Object System.Collections.Generic.List[object]
for ($r = 0; $r -lt [Math]::Min(10, $rkW.rank.Count); $r++) {
  $it = $rkW.rank[$r]
  $mv = if ($prevPos.ContainsKey($it.z)) { $prevPos[$it.z] - ($r + 1) } else { $null }
  $ev = New-Object System.Collections.Generic.List[object]
  foreach ($i in ($it.ev | Select-Object -First 3)) { $ev.Add(@($ALLM[$i].n, (Get-Snip $ALLM[$i].x 90))) }
  $picks.Add([ordered]@{ zone = $it.z; people = $it.people.Count; mentions = $it.men; isNew = (-not $prevPos.ContainsKey($it.z)); move = $mv; evidence = $ev; chats = (Get-Chats $wS $wE $it.z $null $null) })
}
Write-Host ("추천 구역 {0}곳, 부정 소식 구역 {1}곳" -f $picks.Count, $rkW.neg.Count)

# ───────────────────────── 7. 주요 언급 이슈 ─────────────────────────
$topics = @(
  @{ key = '시세'; emoji = '📈'; name = '시세·거래'; re = [regex]'신고가|호가|급매|매물|거래량|실거래|시세|상승|폭등|급등|매수세|프리미엄|피가' },
  @{ key = '하락'; emoji = '📉'; name = '하락장 우려'; re = [regex]'하락장|하락기|하락세|하락할|하락하|폭락|급락|조정장|조정이|조정받|약세|침체|거품|버블|고점|꼭지|물렸|물린|손절|무너지|무너졌|빠지(고|는|네|면|겠)|빠졌|떨어지(고|는|네|면|겠)|떨어졌|하방|공포|역전세|미분양' },
  @{ key = '사업'; emoji = '🏗️'; name = '사업 진행'; re = [regex]'조합설립|조설|추진위|사업시행|사시|관리처분|관처|감정평가|감평|이주|철거|착공|일반분양|일분|동호수|비례율|분담금|시공사|총회|입찰|조합장' },
  @{ key = '정책'; emoji = '📜'; name = '정책·규제'; re = [regex]'토허|토지거래허가|대책|규제|DSR|LTV|대출|스트레스|조정대상|투기과열|실거주\s?의무|양도세|취득세|보유세|종부세|재초환|초과이익|지위양도|전매제한|금리' },
  @{ key = '제도'; emoji = '🏛️'; name = '정비 제도'; re = [regex]'신속통합|신통|모아타운|모아주택|공공재개발|도심공공|정비계획|용적률|기부채납|사업성\s?보정|특례|역세권|정비구역\s?지정|구역지정|후보지|공모|지하화' }
)
function Get-Trend([int]$now, [int]$prev) {
  if ($prev -eq 0) { if ($now -gt 0) { return @('up', '이번 주 새로 늘었어요') } else { return @('', '') } }
  $d = ($now - $prev) / $prev
  if ($d -ge 0.1) { @('up', '지난주보다 늘었어요') } elseif ($d -le -0.1) { @('down', '지난주보다 줄었어요') } else { @('', '지난주와 비슷해요') }
}
$issues = New-Object System.Collections.Generic.List[object]
foreach ($zk in $rkW.neg) {
  $rows = Get-Chats $wS $wE $zk $null $null
  $negMsg = @(for ($i = (Get-LB $wS); $i -lt $ALLM.Count; $i++) { if (($ALLM[$i].z -contains $zk) -and $ALLM[$i].negev) { $ALLM[$i] } }) | Sort-Object { $_.x.Length } -Descending | Select-Object -First 1
  $issues.Add([ordered]@{ key = 'neg:' + $zk; zone = $zk; cat = '🚫 구역 소식'; title = ($zk + ', 부정적인 소식이 이어졌어요'); summary = $(if ($negMsg) { Get-Snip $negMsg.x 80 } else { '' }); people = (Get-People $rows); trend = 'up'; trendText = '이번 주 나온 소식'; chats = $rows })
}
$PASTE = [regex]'운세|일진|음력|조합원 여러분|안내\]|^\s*[#\[【]|공지사항'
$topicCards = foreach ($tp in $topics) {
  $nowP = @{}; $prevP = @{}; $cand = New-Object System.Collections.Generic.List[int]
  for ($i = (Get-LB $pS); $i -lt $ALLM.Count; $i++) {
    $m = $ALLM[$i]; if (-not $tp.re.IsMatch($m.x)) { continue }
    if ($PASTE.IsMatch($m.x)) { continue }   # 운세·공지문 같은 붙여넣기 글은 제외
    if ($m.t -ge $wS) { $nowP[$m.n] = 1; if ($m.x.Length -ge 25 -and $m.x.Length -le 300 -and $m.x -notmatch 'https?://') { $cand.Add($i) } } else { $prevP[$m.n] = 1 }
  }
  $top = @($cand | Sort-Object { $ALLM[$_].x.Length } -Descending | Select-Object -First 12 | Sort-Object)
  $rows = New-Object System.Collections.Generic.List[object]; foreach ($i in $top) { $rows.Add((Get-Row $i 'k')) }
  $byLen = @($top | Sort-Object { $ALLM[$_].x.Length } -Descending)
  $tr = Get-Trend $nowP.Count $prevP.Count
  [ordered]@{ key = $tp.key; cat = ($tp.emoji + ' ' + $tp.name); title = $(if ($byLen.Count) { '“' + (Get-Snip $ALLM[$byLen[0]].x 38) + '”' } else { $tp.name }); summary = $(if ($byLen.Count -gt 1) { Get-Snip $ALLM[$byLen[1]].x 70 } else { '' }); people = $nowP.Count; trend = $tr[0]; trendText = $tr[1]; chats = $rows }
}
foreach ($c in ($topicCards | Sort-Object { $_.people } -Descending)) { if ($c.people -gt 0) { $issues.Add($c) } }

# 갑자기 많이 언급된 구역 (이번 주 ÷ 직전 4주 주평균)
$c4 = @{}; for ($i = (Get-LB $wS.AddDays(-28)); $i -lt $ALLM.Count -and $ALLM[$i].t -lt $wS; $i++) { foreach ($k in $ALLM[$i].z) { $c4[$k] = 1 + [int]$c4[$k] } }
$rising = New-Object System.Collections.Generic.List[object]
$risers = @($statW.Values | Where-Object { $_.men -ge 5 -and $rkW.neg -notcontains $_.z } | ForEach-Object {
  $avg = [double]([int]$c4[$_.z]) / 4
  $ratio = if ($avg -gt 0) { $_.men / $avg } else { 999 }
  [pscustomobject]@{ z = $_.z; men = $_.men; avg = $avg; ratio = $ratio } } | Where-Object { $_.ratio -ge 2 } | Sort-Object ratio -Descending | Select-Object -First 6)
foreach ($r in $risers) {
  $longest = @(for ($i = (Get-LB $wS); $i -lt $ALLM.Count; $i++) { if ($ALLM[$i].z -contains $r.z -and $ALLM[$i].x.Length -ge 12) { $ALLM[$i] } }) | Sort-Object { $_.x.Length } -Descending | Select-Object -First 1
  $rising.Add([ordered]@{ zone = $r.z; count = $r.men; avg = [Math]::Round($r.avg, 1); ratio = $(if ($r.ratio -ge 999) { $null } else { [Math]::Round($r.ratio, 1) }); comment = $(if ($longest) { Get-Snip $longest.x 80 } else { '' }); chats = (Get-Chats $wS $wE $r.z $null $null) })
}

# ───────────────────────── 8. 리트리 ─────────────────────────
$POSRT  = [regex]'(?i)추천|강추|한\s?표|원픽|좋게\s?봅|좋아\s?보|괜찮아\s?보|사세요|사두|담으세요|지켜보시길|지켜\s?보세요|관심\s?가져|저평가|가성비|안전마진|제2의|급이라|꿀매물|좋은\s?매물|사업성\s?나오|주목'
$NEGRT  = [regex]'비추|별로|위험|조심|노답|마피|손절|사지\s?마|취소|무산|해제'
$DANJI  = [regex]'(?<![가-힣])([가-힣]{2,6})\s?(\d{1,2})\s?단지'
$CITE   = [regex]'리트리\s?님?(이|께서|의)?\s?\S{0,6}\s?(추천|픽|찍어|말씀하신)'
function Get-Targets($m) { $t = New-Object System.Collections.Generic.List[string]; foreach ($k in $m.z) { $t.Add($k) }; foreach ($d in $DANJI.Matches($m.x)) { $t.Add($d.Groups[1].Value + $d.Groups[2].Value + '단지') }; ,$t }
function Join-Title($targets) {
  if ($targets.Count -eq 1) { return $targets[0] }
  $p0 = [regex]::Match($targets[0], '^(.+?)(\d.*)$')
  $same = $p0.Success -and @($targets | Where-Object { [regex]::Match($_, '^(.+?)\d').Groups[1].Value -ne $p0.Groups[1].Value }).Count -eq 0
  if ($same) { $p0.Groups[1].Value + (($targets | ForEach-Object { [regex]::Match($_, '\d.*$').Value }) -join '·') } else { $targets -join '·' }
}
function Get-RtPicks([datetime]$a, [datetime]$b) {
  $out = [ordered]@{}
  for ($i = (Get-LB $a); $i -lt $ALLM.Count -and $ALLM[$i].t -lt $b; $i++) {
    $m = $ALLM[$i]
    if ($m.rt) {
      if (-not $POSRT.IsMatch($m.x) -or $NEGRT.IsMatch($m.x) -or $m.q) { continue }
      $tg = Get-Targets $m; $quote = $m.x
      if ($tg.Count -eq 0) {
        for ($j = $i + 1; $j -lt $ALLM.Count -and ($ALLM[$j].t - $m.t).TotalMinutes -le 3; $j++) {
          if (-not $ALLM[$j].rt) { continue }
          $t2 = Get-Targets $ALLM[$j]; if ($t2.Count) { $tg = $t2; $quote = $m.x + ' ' + $ALLM[$j].x; break }
        }
      }
      if ($tg.Count -eq 0) { continue }
      $title = Join-Title $tg
      $out[$title] = [ordered]@{ title = $title; targets = @($tg); date = $m.t.ToString('MM.dd HH:mm'); day = $m.t.ToString('yyyy-MM-dd'); quote = (Get-Snip $quote 160); via = 'direct' }
    } elseif ($CITE.IsMatch($m.x) -and $m.z.Count) {
      $title = Join-Title $m.z
      if (-not $out.Contains($title)) { $out[$title] = [ordered]@{ title = $title; targets = @($m.z); date = $m.t.ToString('MM.dd HH:mm'); day = $m.t.ToString('yyyy-MM-dd'); quote = (Get-Snip $m.x 160); via = 'cite'; by = $m.n } }
    }
  }
  ,@($out.Values | Sort-Object { $_.day } -Descending)
}
$rtWeek = Get-RtPicks $wS $wE
$rtPast = Get-RtPicks $wS.AddDays(-21) $wE
$rtZones = @($rtPast | ForEach-Object { $_.targets })
# 리트리 한마디: 5분 안에 이어서 한 말을 묶고, 부동산 얘기만 남겨 내용 많은 순 2개
$RE_TALK = [regex]'재개발|재건축|정비|조합|사업|매물|분양|입지|지하화|투자|부동산|구역|감평|비례율|분담금|시세|호가|갈아타|역세권|용적률|공시|세금|대출|규제'
$threads = New-Object System.Collections.Generic.List[object]; $cur = $null
for ($i = (Get-LB $wS); $i -lt $ALLM.Count; $i++) {
  $m = $ALLM[$i]; if (-not $m.rt) { continue }
  if ($NOISE.IsMatch($m.x.Trim()) -or $m.x -match '^https?://\S+$') { continue }
  if ($cur -and ($m.t - $cur.last).TotalMinutes -le 5) { $cur.parts.Add($m.x); $cur.last = $m.t }
  else { $cur = [pscustomobject]@{ start = $m.t; last = $m.t; parts = (New-Object System.Collections.Generic.List[string]) }; $cur.parts.Add($m.x); $threads.Add($cur) }
}
$talk = @($threads | ForEach-Object { $txt = ($_.parts -join ' '); [pscustomobject]@{ t = $_.start; n = $_.parts.Count; parts = $_.parts; text = $txt; len = ($txt -replace '[ㅋㅎㅠ^;.!?~\s]', '').Length } } |
  Where-Object { $_.len -ge 40 -and $RE_TALK.IsMatch($_.text) } | Sort-Object len -Descending | Select-Object -First 2 |
  ForEach-Object {
    # 이어지는 문장끼리 한 불렛으로 묶기: "~이자, ~때문에, ~해서"처럼 문장이 덜 끝난 메시지는 다음 메시지와 합침
    $lines = New-Object System.Collections.Generic.List[string]; $buf = ''
    foreach ($p in $_.parts) {
      $s = ($p -replace '\s+', ' ').Trim()
      if (($s -replace '[ㅋㅎㅠ^;.!?~\s]', '').Length -lt 2) { continue }
      $buf = if ($buf) { $buf + ' ' + $s } else { $s }
      $open = $s -match '(,|이자|때문에|지만|는데|은데|인데|해서|어서|라서|이라|으로|에서|하고|이고|면서|거나|든지|및|와|과|랑|고|며|면|니|서|\.\.)$' -or ($s -replace '[ㅋㅎㅠ^;.!?~\s]', '').Length -lt 8
      if (-not $open) { $lines.Add((Get-Snip $buf 200)); $buf = '' }
    }
    if ($buf) { $lines.Add((Get-Snip $buf 200)) }
    $more = [Math]::Max(0, $lines.Count - 6)
    [ordered]@{ lines = @($lines | Select-Object -First 6); more = $more; date = $_.t.ToString('MM.dd HH:mm'); count = $_.n }
  })

# ───────────────────────── 9. 장미마을 ─────────────────────────
$ROSE = [regex]'(?<![가-힣])(신길\s?|도림\s?|영등포\s?)?장미(마을)?'
$ROSEX = [regex]'잠실|하계|상가|잠5|잠주|선수촌|토니안|아크로|자이|은마|여의도'
$mS = [datetime]::new($wS.AddDays(3).Year, $wS.AddDays(3).Month, 1)
$roseW = Get-Chats $wS $wE '' $ROSE $ROSEX
$roseP = Get-Chats $pS $wS '' $ROSE $ROSEX
$roseM = Get-Chats $mS $wE '' $ROSE $ROSEX
$roseTop = @($roseW | Where-Object { $_[3] -eq 'k' } | Sort-Object { $_[2].Length } -Descending | Select-Object -First 3 | ForEach-Object { ,@($_[0], (Get-Snip $_[2] 90)) })

# ───────────────────────── 10. 머리글 ─────────────────────────
$wCount = 0; for ($i = (Get-LB $wS); $i -lt $ALLM.Count -and $ALLM[$i].t -lt $wE; $i++) { $wCount++ }
$latest = $ALLM[$ALLM.Count - 1].t
$head = $null
if ($picks.Count) { $p0 = $picks[0]; $head = [ordered]@{ zone = $p0.zone; people = $p0.people; evidence = $p0.evidence } }

$out = [ordered]@{
  built = (Get-Date).ToString('yyyy-MM-dd HH:mm')
  weekStart = $wS.ToString('yyyy-MM-dd'); weekLabel = (Get-WeekLabel $wS)
  range = ('{0:MM.dd} – {1:MM.dd}' -f $wS, $wE.AddDays(-1)); latest = $latest.ToString('MM.dd HH:mm')
  live = ($latest -lt $wE.AddMinutes(-1)); messages = $wCount; monthLabel = ('{0}월' -f $mS.Month)
  headline = $head
  issues = $issues; rising = $rising
  picks = $picks; negZones = @($rkW.neg)
  retree = [ordered]@{ week = $rtWeek; past = $rtPast; zones = $rtZones; talk = $talk }
  rose = [ordered]@{ week = (Get-KwCount $roseW); prev = (Get-KwCount $roseP); month = (Get-KwCount $roseM); top = $roseTop; chatsWeek = $roseW; chatsMonth = $roseM }
}

# ───────────────────────── 11. weekly.js 쓰기 ─────────────────────────
function ConvertTo-Js($o) {
  if ($null -eq $o) { return 'null' }
  if ($o -is [string]) { return [Web.HttpUtility]::JavaScriptStringEncode($o, $true) }
  if ($o -is [bool]) { if ($o) { return 'true' } else { return 'false' } }
  if ($o -is [int] -or $o -is [long] -or $o -is [double] -or $o -is [decimal]) { return [string]::Format([Globalization.CultureInfo]::InvariantCulture, '{0}', $o) }
  if ($o -is [Collections.IDictionary]) { $parts = foreach ($k in $o.Keys) { (ConvertTo-Js ([string]$k)) + ':' + (ConvertTo-Js $o[$k]) }; return '{' + ($parts -join ',') + '}' }
  if ($o -is [Collections.IEnumerable]) { $parts = foreach ($v in $o) { ConvertTo-Js $v }; return '[' + ($parts -join ',') + ']' }
  return ConvertTo-Js ([string]$o)
}
$js = "window.WEEKLY=" + (ConvertTo-Js $out) + ";`n"
[IO.File]::WriteAllText((Join-Path $root 'weekly.js'), $js, (New-Object Text.UTF8Encoding($false)))
Write-Host ("weekly.js 완료 · {0} ({1}) · {2:N0} KB · {3:N1}초" -f (Get-WeekLabel $wS), $out.range, ($js.Length / 1KB), $clock.Elapsed.TotalSeconds)

if ($args -notcontains '-noopen') { Start-Process (Join-Path $root 'index.html') }
