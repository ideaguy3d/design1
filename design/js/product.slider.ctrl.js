/**
 * Created by user on 4/14/2017.
 */

(function () {
    var app = angular.module('ap-slider');
    app.controller('apSliderCtrl', ['$scope', 'jProductGroup1Data',
        function ($scope, jProductGroup1Data) {
            $scope.productsGroup1_title = "Anzac Day";
            $scope.anzacProducts = jProductGroup1Data.AnzacDayProducts;
            $scope.apcCurrentProducts = jProductGroup1Data.AnzacDayProducts; // will change to different product group later
            $scope.activeArea = -1;
            $scope.repetitionAmount = [0, 1, 2];
            $scope.incrementLeft = false;

            $scope.updateActiveArea = function (index) {
                switch (index) {
                    case 0:
                        $scope.activeArea = -1;
                        break;
                    case 1:
                        $scope.activeArea = 0;
                        break;
                    case 2:
                        $scope.activeArea = 1;
                        break;
                    case 3:
                        $scope.activeArea = 2;
                        break;
                }

                // give the DOM a moment to update
                setTimeout(function(){
                    // pg = product group
                    var pgElem = ".product-group" + (index-1),
                        pgSelector = angular.element(pgElem),
                        pgActive = pgSelector.hasClass('active');

                    if (pgActive) {
                        if(pgElem !== ".product-group2") {
                            console.log("adding increment-left class");
                            pgSelector.addClass('increment-left');
                        } else if (pgElem === ".product-group2") {
                            angular.element('.page-group').removeClass('increment-left');
                            console.log("should be removing increment-left");
                        }
                        console.log("1) jha - active = "+pgElem);
                    } else {
                        console.log("0) jha - active = "+pgElem);
                    }
                }, 50);

                // console.log("jha - activeArea = " + $scope.activeArea + ", incrementLeft = " + $scope.incrementLeft);
            };

            $scope.indicatorCheck = function () {
                var aa = $scope.activeArea;
                if (aa === -1) {
                    return 0;
                } else if (aa === 0) {
                    return 1;
                } else if (aa === 1) {
                    return 2;
                } else if (aa === 2) {
                    return 3;
                }
            };

            var slideProductGroup = function (index) {
                // console.log("jha - slideProductGroup function, do something!");
            };
        }
    ]);
})();

//