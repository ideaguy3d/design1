/**
 * Created by user on 4/28/2017.
 */

(function () {
    "use strict";
    var app = angular.module('ap-slider');
    app.controller('apSliderEditCtrl', ['$scope', '$timeout', 'jProductGroup1Data',
        function ($scope, $timeout, jProductGroup1Data) {
            $scope.productsGroup1_title = "Anzac Day";
            $scope.apcCurrentProducts = jProductGroup1Data.Row1; // will change to different product group later
            $scope.activeArea = -1;
            $scope.repetitionAmount = [0, 1, 2];
            $scope.incrementLeft = false;
            $scope.showImageUrl = true;
            $scope.showHeader = true;
            $scope.showProductId = true;
            $scope.showPrice = true;
            $scope.ratioPortrait = true; // have a function determine if image is portrait
            $scope.ratioLandscape = true; // have a function determine if image is landscape

            //-- private functions:
            var pgPositionTracker = [
                {side: "initialState"},
                {side: "right"},
                {side: "right"},
                {side: "right"}
            ];

            var removeIncrementLeft = function () {
                angular.element('.product-group.increment-left').each(function (i) {
                    // console.log(this);
                    angular.element(this).removeClass('increment-left');
                    pgPositionTracker[i].side = "right";
                });
            };

            var domUpP1 = function (index) {
                // give the DOM a moment to update
                $timeout(function () {
                    // pg = product group
                    var pgElem = ".product-group" + (index - 1),
                        pgSelector = angular.element(pgElem),
                        pgActive = pgSelector.hasClass('active');
                    if (pgActive) {
                        if (pgElem !== '.product-group2') {
                            //'.product-group2' should never increment-left because it's last
                            pgSelector.addClass('increment-left');
                        }
                    }
                }, 10);
            };

            //-- public functions:
            $scope.updateActiveArea = function (index) {
                switch (index) {
                    case 0:
                        pgPositionTracker[0].side = "left";
                        if (pgPositionTracker[1].side === "left" || pgPositionTracker[2].side === "left") {
                            removeIncrementLeft();
                            pgPositionTracker[1].side = "right";
                            pgPositionTracker[2].side = "right";
                        }
                        $scope.activeArea = -1;
                        break;
                    case 1:
                        pgPositionTracker[1].side = "left";
                        if (pgPositionTracker[2].side === "left") {
                            removeIncrementLeft();
                            pgPositionTracker[2].side = "right";
                        }
                        $scope.activeArea = 0;
                        break;
                    case 2:
                        pgPositionTracker[2].side = "left";
                        $scope.activeArea = 1;
                        break;
                    case 3:
                        pgPositionTracker[3].side = "left";
                        $scope.activeArea = 2;
                        // domUpP2();
                        break;
                }

                domUpP1(index);

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

            $scope.editHeader = function () {
                $scope.showPrice = true;
                $scope.showProductId = true;
                $scope.showImageUrl = true;
                $scope.showHeader = !$scope.showHeader;
            };

            $scope.editProductId = function () {
                $scope.showHeader = true;
                $scope.showPrice = true;
                $scope.showImageUrl = true;
                $scope.showProductId = !$scope.showProductId;
            };

            $scope.editPrice = function () {
                $scope.showHeader = true;
                $scope.showProductId = true;
                $scope.showImageUrl = true;
                $scope.showPrice = !$scope.showPrice;
            };

            $scope.editImageUrl = function () {
                $scope.showHeader = true;
                $scope.showProductId = true;
                $scope.showPrice = true;
                $scope.showImageUrl = !$scope.showImageUrl;
            };

            $scope.updateImageUrl = function(product){
                console.log("productId = "+product.image);
                $scope.apcCurrentProducts.$save(product);
            };

            var _newImageUrl_g1p1 = 'Paste image url';

            $scope.row1products = {
                group1product1: {
                    image: function (newImageUrl) {
                        if(newImageUrl === 'g1p1') {
                            return 'group1Product1 imageUrl'
                        } else if (newImageUrl === 0) {
                            _newImageUrl_g1p1 = "John West";
                            return _newImageUrl_g1p1;
                        }

                        // if there isn't an argument this is a getter
                        return arguments.length ? (_newImageUrl_g1p1 = newImageUrl)
                            : _newImageUrl_g1p1;
                    }
                },
                group1product2: {
                    image: function (newImageUrl) {

                    }
                },
                group1product3: {
                    image: function (newImageUrl) {

                    }
                },
                group1product4: {
                    image: function (newImageUrl) {

                    }
                }
            };

            $scope.setRow1 = function(index, product){
                console.log("index = "+index+", product.image = "+product.image);
            };

            $scope.bindToProductsModel = function(index){
              return $scope.g1p1 = "a cool binding";
            };

            $scope.getCategories = "categories from product.slider.edit.ctrl.js file";
        }
    ]);
})();