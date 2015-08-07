	
			var chart;
			$(document).ready(function() {
				chart = new Highcharts.Chart({
					chart: {
						renderTo: 'graph2',
						type: 'spline'
					},
					title: {
						text: 'Number of followers'
					},
					subtitle: {
						text: 'Per Month'	
					},
					xAxis: {
						type: 'datetime',
						dateTimeLabelFormats: { // don't display the dummy year
							month: '%e. %b',
							year: '%b'
						}
					},
					yAxis: {
						title: {
							text: 'Date'
						},
						min: 0
					},
					tooltip: {
						formatter: function() {
				                return '<b>'+ this.series.name +'</b><br/>'+
								Highcharts.dateFormat('%e. %b', this.x) +': '+ this.y +' m';
						}
					},
					
					series: [{
						name: 'Facebook',
						// Define the data points. All series have a dummy year
						// of 1970/71 in order to be compared on the same x axis. Note
						// that in JavaScript, months start at 0 for January, 1 for February etc.
						data: [
							[Date.UTC(1970,  9, 27), 0   ],
							[Date.UTC(1970, 10, 10), 0.6 ],
							[Date.UTC(1970, 10, 18), 0.7 ],
							[Date.UTC(1970, 11,  2), 0.8 ],
							[Date.UTC(1970, 11,  9), 0.6 ],
							[Date.UTC(1970, 11, 16), 0.6 ],
							[Date.UTC(1970, 11, 28), 0.67],
							[Date.UTC(1971,  0,  1), 0.81],
							[Date.UTC(1971,  0,  8), 0.78],
							[Date.UTC(1971,  0, 12), 0.98],
							[Date.UTC(1971,  0, 27), 1.84],
							[Date.UTC(1971,  1, 10), 1.80],
							[Date.UTC(1971,  1, 18), 1.80],
							[Date.UTC(1971,  1, 24), 1.92],
							[Date.UTC(1971,  2,  4), 2.49],
							[Date.UTC(1971,  2, 11), 2.79],
							[Date.UTC(1971,  2, 15), 2.73],
							[Date.UTC(1971,  2, 25), 2.61],
							[Date.UTC(1971,  3,  2), 2.76],
							[Date.UTC(1971,  3,  6), 2.82],
							[Date.UTC(1971,  3, 13), 2.8 ],
							[Date.UTC(1971,  4,  3), 2.1 ],
							[Date.UTC(1971,  4, 26), 1.1 ],
							[Date.UTC(1971,  5,  9), 0.25],
							[Date.UTC(1971,  5, 12), 0   ]
						]
					}, {
						name: 'Youtube',
						data: [
							[Date.UTC(1970,  9, 18), 0   ],
							[Date.UTC(1970,  9, 26), 0.2 ],
							[Date.UTC(1970, 11,  1), 0.47],
							[Date.UTC(1970, 11, 11), 0.55],
							[Date.UTC(1970, 11, 25), 1.38],
							[Date.UTC(1971,  0,  8), 1.38],
							[Date.UTC(1971,  0, 15), 1.38],
							[Date.UTC(1971,  1,  1), 1.38],
							[Date.UTC(1971,  1,  8), 1.48],
							[Date.UTC(1971,  1, 21), 1.5 ],
							[Date.UTC(1971,  2, 12), 1.89],
							[Date.UTC(1971,  2, 25), 2.0 ],
							[Date.UTC(1971,  3,  4), 1.94],
							[Date.UTC(1971,  3,  9), 1.91],
							[Date.UTC(1971,  3, 13), 1.75],
							[Date.UTC(1971,  3, 19), 1.6 ],
							[Date.UTC(1971,  4, 25), 0.6 ],
							[Date.UTC(1971,  4, 31), 0.35],
							[Date.UTC(1971,  5,  7), 0   ]
						]
					}, {
						name: 'Google',
						data: [
							[Date.UTC(1970,  9,  9), 0   ],
							[Date.UTC(1970,  9, 5), 0.15],
							[Date.UTC(1970, 9, 28), 0.35],
							[Date.UTC(1970, 11, 12), 0.46],
							[Date.UTC(1971,  0,  1), 0.59],
							[Date.UTC(1971,  0, 10), 0.58],
							[Date.UTC(1971,  1,  1), 0.62],
							[Date.UTC(1971,  1,  7), 0.65],
							[Date.UTC(1971,  1, 23), 0.77],
							[Date.UTC(1971,  1,  8), 0.77],
							[Date.UTC(1971,  2, 10), 0.79],
							[Date.UTC(1971,  2, 24), 0.86],
							[Date.UTC(1971,  5,  4), 0.8 ],
							[Date.UTC(1971,  3, 30), 0.94],
							[Date.UTC(1971,  6, 22), 0.9 ],
							[Date.UTC(1971,  7, 20), 0.39],
							[Date.UTC(1971,  4, 21), 0   ]
						]
					
					}, {
						name: 'Twitter',
						data: [
							[Date.UTC(1970,  9,  5), 0   ],
							[Date.UTC(1970,  9, 11), 0.15],
							[Date.UTC(1970, 10, 20), 0.35],
							[Date.UTC(1970, 15, 12), 0.46],
							[Date.UTC(1971,  0,  5), 0.59],
							[Date.UTC(1971,  0, 11), 0.58],
							[Date.UTC(1971,  1,  10), 0.62],
							[Date.UTC(1971,  1,  5), 0.65],
							[Date.UTC(1971,  2, 12), 0.77],
							[Date.UTC(1971,  2,  5), 0.77],
							[Date.UTC(1971,  2, 10), 0.79],
							[Date.UTC(1971,  2, 22), 0.86],
							[Date.UTC(1971,  3,  5), 0.8 ],
							[Date.UTC(1971,  3, 2), 0.94],
							[Date.UTC(1971,  3, 14), 0.9 ],
							[Date.UTC(1971,  4, 16), 0.39],
							[Date.UTC(1971,  4, 21), 0   ]
						]
					
					}, {
						name: 'Linkedin',
						data: [
							[Date.UTC(1970,  9,  9), 0   ],
							[Date.UTC(1970,  9, 14), 0.35],
							[Date.UTC(1970, 10, 28), 0.45],
							[Date.UTC(1970, 11, 12), 0.46],
							[Date.UTC(1971,  0,  1), 0.59],
							[Date.UTC(1971,  0, 24), 0.78],
							[Date.UTC(1971,  1,  1), 0.62],
							[Date.UTC(1971,  1,  7), 0.25],
							[Date.UTC(1971,  1, 23), 0.67],
							[Date.UTC(1971,  2,  8), 0.37],
							[Date.UTC(1971,  2, 14), 0.79],
							[Date.UTC(1971,  2, 24), 0.76],
							[Date.UTC(1971,  3,  4), 0.1 ],
							[Date.UTC(1971,  3, 18), 0.34],
							[Date.UTC(1971,  3, 24), 0.9 ],
							[Date.UTC(1971,  4, 16), 0.39],
							[Date.UTC(1971,  4, 21), 0   ]
						]
						
					}, {
						name: 'Overall',
						data: [
							[Date.UTC(1970,  9,  9), 0   ],
							[Date.UTC(1970,  9, 14), 0.85],
							[Date.UTC(1970, 10, 11), 0.35],
							[Date.UTC(1970, 11, 18), 0.16],
							[Date.UTC(1971,  0,  1), 0.59],
							[Date.UTC(1971,  0, 24), 0.88],
							[Date.UTC(1971,  1,  1), 0.12],
							[Date.UTC(1971,  1,  7), 0.65],
							[Date.UTC(1971,  1, 18), 0.77],
							[Date.UTC(1971,  2,  8), 0.87],
							[Date.UTC(1971,  2, 14), 0.79],
							[Date.UTC(1971,  2, 14), 0.86],
							[Date.UTC(1971,  3,  8), 0.8 ],
							[Date.UTC(1971,  3, 11), 0.84],
							[Date.UTC(1971,  3, 14), 0.9 ],
							[Date.UTC(1971,  4, 16), 0.39],
							[Date.UTC(1971,  4, 11), 0   ]
						]
						
					}]
				});
				
				
			});
				
