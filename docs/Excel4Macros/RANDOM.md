RANDOM
======

Fills a range with independent random or patterned numbers drawn from
one of several distributions.

If this function is not available, you must install the Analysis ToolPak
add-in.

RANDOM provides six different random distributions and one patterned
data option. Because the distributions require different argument lists,
there are seven syntax forms for RANDOM.

**Syntax 1**

Uniform distribution

**RANDOM**(outrng, variables, points, **distribution**, seed, **from,
to**)

**RANDOM**?(outrng, variables, points, distribution, seed, from, to)

**Syntax 2**

Normal distribution

**RANDOM**(outrng, variables, points, **distribution**, seed, **mean,
standard\_dev**)

**RANDOM**?(outrng, variables, points, distribution, seed, mean,
standard\_dev)

**Syntax 3**

Bernoulli distribution

**RANDOM**(outrng, variables, points, **distribution**, seed,
**probability**)

**RANDOM**?(outrng, variables, points, distribution, seed, probability)

**Syntax 4**

Binomial distribution

**RANDOM**(outrng, variables, points, **distribution**, seed,
**probability, trials**)

**RANDOM**?(outrng, variables, points, distribution, seed, probability,
trials)

**Syntax 5**

Poisson distribution

**RANDOM**(outrng, variables, points, **distribution**, seed,
**lambda**)

**RANDOM**?(outrng, variables, points, distribution, seed, lambda)

**Syntax 6**

Patterned distribution

**RANDOM**(outrng, variables, points, **distribution**, seed, **from,
to, step, repeat\_num, repeat\_seq**)

**RANDOM**?(outrng, variables, points, distribution, seed, from, to,
step, repeat\_num, repeat\_seq)

**Syntax 7**

Discrete distribution

**RANDOM**(outrng, variables, points, **distribution**, seed,
**inprng**)

**RANDOM**?(outrng, variables, points, distribution, seed, inprng)

Outrng    is the first cell (the upper-left cell) in the output table or
the name, as text, of a new sheet to contain the output table. If FALSE,
blank, or omitted, places the output table in a new workbook.

Variables    is the number of random number sets to generate. RANDOM
will generate variables columns of random numbers. If omitted, variables
is equal to the number of columns in the output range.

Points    is the number of data points per random number set. RANDOM
will generate points rows of random numbers for each random number set.
If omitted, points is equal to the number of rows in the output range.
Points is ignored when distribution is 6 (Patterned).

Distribution    indicates the type of number distribution.

+--------------------+-------------------------+
| > **Distribution** | > **Distribution type** |
+--------------------+-------------------------+
| > 1                | > Uniform               |
+--------------------+-------------------------+
| > 2                | > Normal                |
+--------------------+-------------------------+
| > 3                | > Bernoulli             |
+--------------------+-------------------------+
| > 4                | > Binomial              |
+--------------------+-------------------------+
| > 5                | > Poisson               |
+--------------------+-------------------------+
| > 6                | > Patterned             |
+--------------------+-------------------------+
| > 7                | > Discrete              |
+--------------------+-------------------------+

Seed    is an optional value with which to begin random number
generation. Seed is ignored when distribution is 6 (Patterned) or 7
(Discrete).

From    is the lower bound.

To    is the upper bound.

Mean    is the mean.

Standard\_dev    is the standard deviation.

Probability    is the probability of success on each trial.

Trials    is the number of trials.

Lambda    is the Poisson distribution parameter.

Step    is the increment between from and to.

Repeat\_num    is the number of times to repeat each value.

Repeat\_seq    is the number of times to repeat each sequence of values.

Inprng    is a two-column range of values and their probabilities.

Return to [top](#Q)

RANKPERC
