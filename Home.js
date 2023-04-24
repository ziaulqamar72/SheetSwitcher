var app = angular.module('myApp', []);
app.controller('myCtrl', function ($scope) {



    Office.onReady(function () {
        var currentActiveSheet;
        var listOfDeactiveSheets = [];

        Excel.run(function (context) {
            const sheet = context.workbook.worksheets;
            sheet.onDeactivated.add(deactiveSheets);
            return context.sync().then(function () {

                //   console.log(currentActiveSheet);
            });
        });
        deactiveSheets();

        function deactiveSheets(event) {
            listOfDeactiveSheets = [];
            $scope.currentDeactiveSheets=[];
         //   console.log(event.worksheetId);
            //   console.log();
            Excel.run(function (context) {
                let activeSheet = context.workbook.worksheets.getActiveWorksheet();
                let sheets = context.workbook.worksheets;
                sheets.load("items/name");
                return context.sync().then(function () {
                    currentActiveSheet = activeSheet.load("name");
                    
                
                    let totelSheets = sheets.items;
                    console.log(totelSheets);
                    for (let i = 0; i < totelSheets.length; i++) {
                        if (totelSheets[i].id != currentActiveSheet.id) {
                            listOfDeactiveSheets.push(totelSheets[i].name)
                        } else {
                            $scope.CurrentActiveSheetName = totelSheets[i].name;
                        }
                        if (!$scope.$$phase) {
                            $scope.$apply();
                        }
                    }
                    console.log(listOfDeactiveSheets);
                    $scope.currentDeactiveSheets = listOfDeactiveSheets;
                    if (!$scope.$$phase) {
                        $scope.$apply();
                    }
                });
            });
        };


    })

});

