$script:AliasInitials = @(
    "j", "m", "a", "s", "d", "k", "r", "t", "l", "c", "b", "n", "e",
    "p", "g", "h", "f", "w", "v", "i", "o", "q", "u", "x", "y", "z"
)

$script:AliasSurnames = @(
    "peterson", "carter", "brooks", "bennett", "collins", "turner", "parker",
    "reed", "morgan", "cooper", "bailey", "murphy", "kelly", "price", "bell",
    "ward", "cook", "gray", "ross", "foster", "powell", "long", "patterson",
    "hughes", "flores", "washington", "butler", "simmons", "fisher", "barnes",
    "henderson", "coleman", "jenkins", "perry", "russell", "griffin", "diaz",
    "hayes", "myers", "ford", "hamilton", "graham", "sullivan", "wallace",
    "woods", "cole", "west", "jordan", "owens", "reynolds", "ellis", "harrison",
    "gibson", "mcdonald", "cruz", "marshall", "ortiz", "gomez", "murray",
    "freeman", "wells", "webb", "simpson", "stevens", "tucker", "porter",
    "hunter", "hicks", "crawford", "henry", "boyd", "mason", "morales",
    "kennedy", "warren", "dixon", "ramos", "reyes", "burns", "gordon",
    "shaw", "holmes", "rice", "robertson", "hunt", "black", "daniels",
    "palmer", "mills", "nichols", "grant", "knight", "ferguson", "rose",
    "stone", "hawkins", "dunn", "perkins", "hudson", "spencer", "gardner"
)

function Get-HumanAliasLocalPart {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateRange(1, 2147483647)]
        [int]$Sequence
    )

    $index = [int64]$Sequence - 1
    $pairCount = [int64]$script:AliasInitials.Count * [int64]$script:AliasSurnames.Count
    $pairIndex = $index % $pairCount

    $initial = $script:AliasInitials[$pairIndex % $script:AliasInitials.Count]
    $surname = $script:AliasSurnames[$pairIndex % $script:AliasSurnames.Count]
    $cycle = [Math]::Floor($index / $pairCount)

    $suffix = if ($cycle -gt 0) {
        ([int64]$cycle + 1).ToString()
    }
    else {
        ""
    }

    return "$initial$surname$suffix"
}
