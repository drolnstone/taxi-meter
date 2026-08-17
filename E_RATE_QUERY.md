# E rate query — ready to send to Liverpool licensing

Copy the block below. It is written so the answer can only be "yes, it's 15" or
"no, it's a typo", with nothing left to interpret. Send it now rather than in
December — a tariff correction has to go through the authority, and that takes time.

---

**Subject: Query on the E rate distance increment — Hackney Carriage Rates of Fare, operative 31 March 2025**

Dear Licensing,

I am building a taximeter application configured against The City of Liverpool
Hackney Carriage Rates of Fare, operative from 31 March 2025, and I need to confirm
one figure before it goes into service.

The E (Extra) rate is printed as:

> £5.10 for the first 300 yards or less
> Then 30p each succeeding **15 yards** (or less) up to 45,000 yards
> Then 30p each succeeding 300 yards (or less)
> And 30p each 60 seconds waiting charge (or less)

Taken literally, 30p per 15 yards is **£35.20 per mile**. For comparison, using the
same document:

| Rate | Increment | Step | Per mile |
|---|---|---|---|
| D (Day) | 20p | 165 yd | £2.13 |
| N (Night) | 25p | 165 yd | £2.67 |
| P (Peak) | 25p | 155 yd | £2.84 |
| **E (Extra), as printed** | **30p** | **15 yd** | **£35.20** |

That is 12.4 times the Peak rate. In fares, a two-mile journey on Christmas Eve
would meter at **£69.60**, and an eight-mile airport run at **£280.80**, against
£7.40 and £20.20 respectively on the Day rate.

Two further points suggest a typographical error rather than an intended rate:

1. Every other band expresses its premium through the increment amount, not the
   step distance. D and N share an identical 165-yard step and differ only in price.
   The E rate's flag (£5.10) is 1.5× the Day flag and its increment (30p) is 1.5×
   the Day increment, both consistent with a seasonal premium; only the step
   distance is out of pattern, and by a factor of eleven.

2. The same sentence provides for the step to widen to 300 yards beyond 45,000
   yards. At 15-yard steps a fare would have reached approximately £899 before that
   threshold, so the provision could never operate. At 165 yards it operates
   normally, as it does for the D and N rates.

Could you please confirm in writing which of the following is the approved figure:

- **(a)** 30p each succeeding **15 yards** — i.e. £35.20 per mile is correct and intended; or
- **(b)** 30p each succeeding **165 yards** — i.e. the published document contains a
  typographical error, matching the D and N rates.

If (b), I would be grateful if the published table could be corrected, as it is the
document drivers and passengers are directed to.

I would also be grateful for confirmation on a second point. The P (Peak) rate is
described as applying to journeys "beginning after 21.00 on a Friday and ending
07.00 on a Saturday, and before 08:00 on Sunday, Easter Sunday, bank holidays apart
from at Christmas and New Year". I have read the second clause as an elliptical
parallel to the first — that is, journeys beginning after 21:00 on a Saturday and
ending before 08:00 on Sunday. Could you confirm whether that is correct, and
whether Easter Sunday and the qualifying bank holidays attract the P rate for the
whole day or only during specified hours.

Many thanks,

[name] — [licence number]

---

## Until it comes back

The meter is set to **165 yd** and says so plainly on its Settings screen, with the
disputed figure and the consequence stated. The constant is a single line:

```js
const E_STEP_YARDS_UNCONFIRMED = 165   /* 165 = reading it as a typo. 15 = as printed. */
```

Changing it to 15 is one character, and I will do it the moment you have (a) in
writing. What I am not willing to do is ship it on the strength of the document
alone, because a driver metering £69.60 for a two-mile journey is not a compliance
question — it is an accusation, and the driver is the one who has to answer it at
the rank, not the person who typed the table.

The E rate first becomes live at **18:00 on 24 December**. That is the deadline.
