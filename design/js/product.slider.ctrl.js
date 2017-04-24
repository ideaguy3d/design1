/**
 * Created by user on 4/14/2017.
 */

(function(){
    var app = angular.module('ap-slider');
    app.controller('apSliderCtrl', ['$scope', 'jProductGroup1Data',
        function($scope, jProductGroup1Data){
            $scope.productsGroup1_title = "Anzac Day Products";
            $scope.anzacProducts = jProductGroup1Data.AnzacDayProducts;
            $scope.activeArea = 0;

            $scope.updateActiveArea = function(index){
                console.log($scope.activeArea+", $index = "+index);
                $scope.activeArea = index;
            }
        }
    ]);
})();

//