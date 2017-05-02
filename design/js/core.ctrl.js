/**
 * Created by user on 4/21/2017.
 */

(function () {
    var app = angular.module('ap-slider', ['firebase']);

    // products group 1 data service
    app.factory("jProductGroup1Data", ['$firebaseObject', '$firebaseArray',
        function ($firebaseObject, $firebaseArray) {
            var ref_productData = firebase.database().ref().child('productData');
            var ref_messages = firebase.database().ref().child('messages');
            var syncObject = $firebaseObject(ref_productData);
            var ref_row1 = firebase.database().ref().child('Row1');

            return {
                Row1: $firebaseArray(ref_row1),
                ProductsMessagesArray: function () {
                    return $firebaseArray(ref_messages);
                }
            }
        }
    ]);

    // apCoreCtrl
    app.controller('apcCoreCtrl', ['$scope', 'jProductGroup1Data',
        function ($scope, jProductGroup1Data) {

            $scope.messages = jProductGroup1Data.ProductsMessagesArray();
            $scope.intro_message = "Ello World ^_^/";
            $scope.showSlider =  location.host === 'www.aussieproducts.com';
            console.log($scope.showSlider);

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