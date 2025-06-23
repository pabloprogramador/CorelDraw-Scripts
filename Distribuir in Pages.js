let shapes = host.ActiveSelectionRange;
let cont = 0;

 alert("Localizado! "+ (shapes.Count - 1) +" objetos");

 let pag = host.ActiveDocument.InsertPagesEx(shapes.Count - 1, false, host.ActivePage.Index, 4.16665, 7.28346);

for (let i = 1; i <= shapes.Count; i++) {
    let shape = shapes.Item(i);
    //shape.CreateSelection();

    //pag.activate(); 

    shape.MoveToLayer(host.ActiveDocument.Pages.Item(i).Layers.Item("Layer 1"));

	cont++;
}

 alert("Processo concluído! "+ cont +" páginas");