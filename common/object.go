package common

type Object map[string]interface{}

func (o Object) Set(key string, val interface{}) {
	if o == nil {
		o = Object{}
	}
	o[key] = val
}
