let OrigSelection = host.ActiveSelectionRange;

for (let i = 1; i <= OrigSelection.Count; i++) {
    OrigSelection.Item(i).Outline.SetNoOutline();
}