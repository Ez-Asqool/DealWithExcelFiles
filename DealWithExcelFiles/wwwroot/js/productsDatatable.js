$(document).ready(function () {
    $('#products').dataTable({
        "serverSide": true,
        "filter": true,
        "ajax": {
            "url": "/Product/AllData",
            "type": "POST",
            "datatype": "JSON"
        },
        "columnDefs": [
            {
                "targets": [0],
                "visible": false,
                "searchable": false
            }
        ],
        "columns": [
            { "data": "id", "name": "Id", "autowidth": true },
            {
                data: 'band',
                name: 'Band',
                autowidth: true,
                render: function (data, type, row) {
                    return '<div style="font-size: 16px;">' + data + '</div>';
                }
            },
            {
                data: 'categoryCode',
                name: 'CategoryCode',
                autowidth: true,
                render: function (data, type, row) {
                    return '<div style="font-size: 16px;">' + data + '</div>';
                }
            },
            {
                data: 'manufacturer',
                name: 'Manufacturer',
                autowidth: true,
                render: function (data, type, row) {
                    return '<div style="font-size: 16px;">' + data + '</div>';
                }
            },
            {
                data: 'partSKU',
                name: 'PartSKU',
                autowidth: true,
                render: function (data, type, row) {
                    return '<div style="font-size: 16px;">' + data + '</div>';
                }
            },
            {
                data: 'itemDescription',
                name: 'ItemDescription',
                autowidth: true,
                render: function (data, type, row) {
                    return '<div style="font-size: 16px;">' + data + '</div>';
                }
            },
            {
                data: 'listPrice',
                name: 'ListPrice',
                autowidth: true,
                render: function (data, type, row) {
                    return '<div style="font-size: 16px;">' + data + '</div>';
                }
            },
            {
                data: 'minimumDiscount',
                name: 'MinimumDiscount',
                autowidth: true,
                render: function (data, type, row) {
                    return '<div style="font-size: 16px;">' + data + '</div>';
                }
            },
            {
                data: 'discountedPrice',
                name: 'DiscountedPrice',
                autowidth: true,
                render: function (data, type, row) {
                    return '<div style="font-size: 16px;">' + data + '</div>';
                }
            }

        ]
    });


    var dataTable = $('#products').DataTable();

    $('#exportButton').click(function () {
        // Get the search value from DataTable's search input
        var searchValue = dataTable.search();

        // Send the search value to the server
        $.ajax({
            url: '/Product/ExportToExcel',
            method: 'POST',
            data: { searchValue: searchValue }, // Pass the search value to the server
            success: function (response) {
                window.location.href = '/Product/DownloadExcel?fileName=' + response.fileName;
                console.log("Exported Successfully");
            },
            error: function () {
                alert('An error occurred while exporting to Excel.');
            }
        });
    });

});


