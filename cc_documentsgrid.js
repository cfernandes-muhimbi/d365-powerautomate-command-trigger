function openPopup(primaryControl) {
    try {
        const formContext = primaryControl || Xrm.Page;

        const gridControl = formContext.getControl("Subgrid_Documents");

        if (!gridControl) {
            alert("SubGrid not found.");
            return;
        }

        const grid = gridControl.getGrid();

        if (!grid) {
            alert("Grid not loaded yet.");
            return;
        }

        // ✅ ONLY SELECTED ROWS
        const rows = grid.getSelectedRows();

        if (!rows || rows.getLength() === 0) {
            alert("Please select at least one document.");
            return;
        }

        let items = [];

        rows.forEach(function (row) {
            const entity = row.getData().getEntity();
            const attr = entity.attributes.getByName("absoluteurl");

            items.push({
                id: entity.getId(),
                name: entity.getPrimaryAttributeValue(),
                absoluteurl: attr ? attr.getValue() : null
            });
        });

        const pageInput = {
            pageType: "webresource",
            webresourceName: "cc_cc_documentpopup1",
            data: JSON.stringify(items)
        };

        const navigationOptions = {
            target: 2,
            width: { value: 600, unit: "px" },
            height: { value: 400, unit: "px" },
            position: 1
        };

        Xrm.Navigation.navigateTo(pageInput, navigationOptions);

    } catch (e) {
        alert("Error: " + e.message);
    }
}