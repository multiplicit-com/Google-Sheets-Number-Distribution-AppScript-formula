function distributeGoal(distributionType, totalMonths, currentPosition, target) {
  // Initialize the value for the current position
  var value = 0;

  switch (distributionType) {
    case 'linear':
      // Linear distribution: equally distribute the goal among the months
      value = target / totalMonths;
      break;

    case 'logarithmic':
      // Logarithmic distribution
      var logSum = 0;
      for (var i = 1; i <= totalMonths; i++) {
        logSum += Math.log(1 + i);
      }
      var logValue = Math.log(1 + currentPosition);
      value = (logValue / logSum) * target;
      break;

    case 'exponential':
      // Exponential distribution
      var expSum = 0;
      for (var i = 1; i <= totalMonths; i++) {
        expSum += Math.exp(i);
      }
      var expValue = Math.exp(currentPosition);
      value = (expValue / expSum) * target;
      break;

    case 'normal':
      // Normal (Gaussian) distribution
      var mean = (totalMonths + 1) / 2;
      var variance = totalMonths / 6;
      var normalSum = 0;
      for (var i = 1; i <= totalMonths; i++) {
        normalSum += Math.exp(-Math.pow(i - mean, 2) / (2 * variance));
      }
      var normalValue = Math.exp(-Math.pow(currentPosition - mean, 2) / (2 * variance));
      value = (normalValue / normalSum) * target;
      break;

    case 'quadratic':
      // Quadratic distribution (Ski Jump style)
      var quadSum = 0;
      for (var i = 1; i <= totalMonths; i++) {
        quadSum += Math.pow(i, 2);
      }
      var quadValue = Math.pow(currentPosition, 2);
      value = (quadValue / quadSum) * target;
      break;

    default:
      throw new Error('Unknown distribution type: ' + distributionType);
  }

  return value;
}

// Custom function callable from the sheet
function DISTRIBUTEGOAL(distributionType, totalMonths, currentPosition, target) {
  return distributeGoal(distributionType, totalMonths, currentPosition, target);
}
