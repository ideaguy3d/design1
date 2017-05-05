/**
 * Created by user on 4/21/2017.
 */

(function () {
    var app = angular.module('ap-slider');

    // apcCoreCtrl
    app.controller('apcCoreCtrl', ['$scope', 'jProductGroup1Data',
        function ($scope, jProductGroup1Data) {
            $scope.messages = jProductGroup1Data.ProductsMessagesArray();
            $scope.intro_message = "Ello World ^_^/";
            $scope.showSlider =  location.host === 'www.aussieproducts.com';

            $scope.addMessage = function () {
                $scope.messages.$add({
                    text: $scope.newMessageText,
                    zdate: Date.now()
                });

                $scope.newMessageText = '';
            };

            //-- 3 way data binding:
            // syncObject.$bindTo($scope, 'data');
        }
    ]);
})();


//