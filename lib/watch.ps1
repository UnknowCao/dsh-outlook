# dsh-outlook NewMailEx watcher (persistent process).
# Spawned by index.js; prints one JSON object per line to STDOUT:
#   { "type": "hello",  "unread": N }              on startup
#   { "type": "newmail", entryId, subject, sender, received }  per new mail
#   { "type": "tick",   "unread": N }              when idle 60s
# Outlook must be running (COM automation attaches to the live client);
# this script waits for it indefinitely.
#
# Event plumbing note (validated on this build, PS 5.1):
#  - Register-ObjectEvent WITHOUT -Action + Wait-Event is the only reliable
#    pattern here: a -Action scriptblock never runs while the main pipeline
#    sits inside Start-Sleep, so events were silently lost.
#  - The COM delegate args arrive as $e.SourceArgs[0] (string of
#    comma-separated EntryIds); $e.SourceEventArgs stays null for COM events.

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function Emit($obj) {
  [Console]::Out.WriteLine(($obj | ConvertTo-Json -Compress -Depth 3))
  [Console]::Out.Flush()
}

function GetOutlook {
  while ($true) {
    try { return [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application') } catch { Start-Sleep -Seconds 5 }
  }
}

$outlook = GetOutlook
$ns = $outlook.GetNamespace('MAPI')
try { Emit @{ type = 'hello'; unread = $ns.GetDefaultFolder(6).UnReadItemCount } } catch { Emit @{ type = 'hello'; unread = $null } }

Register-ObjectEvent -InputObject $outlook -EventName NewMailEx -SourceIdentifier OlkNewMail

while ($true) {
  $e = Wait-Event -SourceIdentifier OlkNewMail -Timeout 60
  if ($e) {
    try { Remove-Event -EventIdentifier $e.EventIdentifier } catch { }
    $ids = [string]@($e.SourceArgs)[0]
    foreach ($id in ($ids -split ',')) {
      $id = $id.Trim()
      if ($id -eq '') { continue }
      try {
        $it = $ns.GetItemFromID($id)
        # Use the item's own ReceivedTime (fall back to now if unavailable).
        $recv = ''
        try { $recv = $it.ReceivedTime.ToString('yyyy-MM-dd HH:mm') } catch { $recv = (Get-Date).ToString('yyyy-MM-dd HH:mm') }
        Emit @{
          type     = 'newmail'
          entryId  = $id
          subject  = [string]$it.Subject
          sender   = [string]$it.SenderName
          received = $recv
        }
      } catch { /* item vanished mid-sync */ }
    }
  } else {
    # Idle heartbeat: the timeout branch also keeps the unread count fresh.
    try { Emit @{ type = 'tick'; unread = $ns.GetDefaultFolder(6).UnReadItemCount } } catch { }
  }
}
