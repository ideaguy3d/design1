/**
 * Created by user on 4/28/2017.
 */

(function () {
    "use strict";
    var app = angular.module('ap-slider');
    app.controller('apSliderEditCtrl', ['$scope', '$timeout', 'jProductGroup1Data',
        function ($scope, $timeout, jProductGroup1Data) {
            $scope.apcRow1Group1 = jProductGroup1Data.Row1Group1;
            $scope.apcRow1Group2 = jProductGroup1Data.Row1Group2;
            $scope.apcRow1Group3 = jProductGroup1Data.Row1Group3;
            $scope.apcRow1Group4 = jProductGroup1Data.Row1Group4;
            $scope.activeArea = -1;
            $scope.repetitionAmount = [0, 1, 2];
            $scope.incrementLeft = false;
            $scope.ratioPortrait = true; // have a function determine if image is portrait
            $scope.ratioLandscape = true; // have a function determine if image is landscape
            $scope.apcProducts = [];
            // temporary data model:
            $scope.exposedRow1Model = {
                image: 'http://www.aussieproducts.com/images/ARTWORK%20IMAGES/Website%20Images/summerroo.png',
                name: 'temp placeholder',
                productId: '12345',
                price: 19.95
            };

            //-- private vars/functions:
            var pgPositionTracker = [
                {side: "initialState"},
                {side: "right"},
                {side: "right"},
                {side: "right"}
            ];

            var indexer = 0, prodGrpArr = [], counter = 0, k;

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

            var _newImageUrl_g1p1 = 'Paste image url';

            $scope.row1products = {
                group1product1: {
                    image: function (newImageUrl) {
                        if (newImageUrl === 'g1p1') {
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

            $scope.setRow1 = function (index, productGroup) {
                // ensure for in loop happens only once
                if (indexer < 1) {
                    // fill the private prodGrpArr
                    for (k in $scope.apcRow1) {
                        // make sure 'k' is NOT being inherited
                        if ($scope.apcRow1.hasOwnProperty(k)) {
                            prodGrpArr[counter] = $scope.apcRow1[k];
                            counter++;
                        }
                    }

                    // copy into $scope.apcProducts only what is needed from prodGrpArr
                    for (var i = 0; i < 4; i++) {
                        $scope.apcProducts[i] = prodGrpArr[i];
                        // get rid of keys that have a '$'
                        var keys = Object.keys($scope.apcProducts[i]);
                        keys = keys.filter(function (e) {
                            return e.includes('$');
                        });
                        // $scope.apcProducts should be cleaned up after this loop
                        keys.forEach(function (e) {
                            delete $scope.apcProducts[i][e];
                        });
                    }
                }
                indexer++;
            };

            $scope.oldSetRow1 = function (index, product) {
                switch (product.$id) {
                    case "Group1":
                        exposeRow1Model(product);
                        break;
                    case "Group2":
                        exposeRow1Model(product);
                        break;
                    case "Group3":
                        exposeRow1Model(product);
                        break;
                    case "Group4":
                        exposeRow1Model(product);
                        break;
                    default:
                        console.log("ERROR: parent index out of range!")
                }
            };

            var exposeRow1Model = function (productGroup) {
                var prodGrpArr = [], counter = 0, k;
                for (k in productGroup) {
                    // make sure 'k' is NOT being inherited
                    if (productGroup.hasOwnProperty(k)) {
                        prodGrpArr[counter] = productGroup[k];
                        counter++;
                    }
                }
                console.log(prodGrpArr);
            };

            $scope.bindToProductsModel = function (index) {
                return $scope.g1p1 = "a cool binding";
            };

            $scope.getCategories = "categories from product.slider.edit.ctrl.js file";
        }
    ]);
})();