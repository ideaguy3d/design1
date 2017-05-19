/**
 * Created by Julius Alvarado on 5/5/2017.
 */

(function () {
    var app = angular.module('ap-slider', ['firebase', 'ngRoute']);

    app.config(['$routeProvider', '$locationProvider',
        function ($routeProvider, $locationProvider) {
            $routeProvider
                .when('/', {
                    templateUrl: 'design/js/views/view.login.html'
                })
                .when('/preview', {
                    templateUrl: 'design/js/views/view.preview.product.slider.html'
                })
                .when('/edit', {
                    templateUrl: 'design/js/views/view.edit.products.html'
                })
                .when('/trialanderror', {
                    templateUrl: 'design/js/views/view.trial.and.error.html'
                });
        }
    ]);
})();

//