const { list:{ operation:{ perform: listBuckets}}} =  require("../../resources/bucket");

const findBucket = (z, bundle, bucket) => listBuckets(z, bundle)
    .then(buckets => 
        buckets.find(
            ({id, name}) => [id, name].includes(bucket)
        )
    )

      
module.exports = { findBucket }
  