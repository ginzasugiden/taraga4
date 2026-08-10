# clasp wrapper for this project.
# Always forces the "tokyoflower" (tokyoflowerco.ltd@gmail.com) clasp credential profile,
# because the machine-wide default clasp profile is bound to a different, unrelated
# Google account (otajigyokyo@gmail.com). See CLAUDE.md for details.
#
# Usage: .\clasp.ps1 <clasp subcommand and args...>
#   e.g. .\clasp.ps1 status
#        .\clasp.ps1 push

param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$Args
)

& clasp -u tokyoflower @Args
