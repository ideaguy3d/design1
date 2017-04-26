/**
 * Created by user on 4/14/2017.
 */

(function(){
    var app = angular.module('ap-slider');
    app.controller('apSliderCtrl', ['$scope', 'jProductGroup1Data',
        function($scope, jProductGroup1Data){
            $scope.productsGroup1_title = "Anzac Day Products";
            $scope.anzacProducts = jProductGroup1Data.AnzacDayProducts;
            $scope.apcCurrentProducts = jProductGroup1Data.AnzacDayProducts; // will change to different product group later
            $scope.activeArea = 0;
            $scope.moveInitialGroup = false;
            $scope.repetitionAmount = [0,1,2];
            $scope.productGroup = -1;

            $scope.updateActiveArea = function(index){
                $scope.productGroup = index;
                console.log($scope.activeArea+", $index = "+index);
                $scope.activeArea = index;
                slideProductGroup(index);
            };

            var slideProductGroup = function(index){
                console.log("jha - slideProductGroup function, do something!");
                $scope.moveInitialGroup = true;
            };
        }
    ]);
})();

//